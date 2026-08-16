import asyncio
import math
import re
from datetime import datetime, timedelta
from typing import Any, Optional

import pandas as pd

from excelmcp.embeddings import search
from excelmcp.graph_client import GraphAPIError, get_used_range
from excelmcp.ranges import parse_span, region_to_df_slice
from excelmcp.storage import log
from excelmcp.structure import (
    columns_from_row,
    fuzzy_name_candidates,
    load_graph,
    normalise_name,
    require_current_schema,
    workspace_relationships,
)

# Hard ceiling on rows returned by a single tool call. An unfiltered sheet can
# hold hundreds of thousands of rows; returning them all would blow out the
# agent's context window and the MCP message size. Callers see `truncated` and
# `total_matched` so they can narrow the query instead of silently losing rows.
MAX_ROWS_RETURNED = 1000
QUERY_ROWS_PER_SHEET = 50


def _records(df: pd.DataFrame) -> list[dict[str, Any]]:
    """Converts a DataFrame to JSON-safe records.

    `to_dict(orient="records")` yields numpy scalars and NaN/NaT, none of which
    survive strict JSON serialisation — NaN in particular encodes as the literal
    `NaN`, which is invalid JSON and broke the MCP response for any sheet with a
    blank cell. Nulls become None and numpy scalars become Python natives.
    """
    if df.empty:
        return []
    safe = df.astype(object).where(pd.notna(df), None)
    records = safe.to_dict(orient="records")
    for row in records:
        for key, value in row.items():
            if hasattr(value, "item"):  # numpy scalar
                value = value.item()
                row[key] = value
            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                row[key] = None
    return records


# Excel serials count days from this epoch (the 1900 system with its
# deliberate Lotus-compat off-by-two).
_EXCEL_EPOCH = datetime(1899, 12, 30)

# Matches an ISO date or datetime literal: 2026-01-01, 2026-01-01T09:30, ...
_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([T ].*)?$")


def _serial_to_iso(value: Any) -> Any:
    """Converts one Excel serial to an ISO-8601 string; non-numerics pass through.

    A whole-day serial becomes a date ('2026-03-15'); a fractional one keeps
    its time part ('2026-03-15T09:30:00'). Values that cannot be a date
    (text, None, out-of-range numbers) are returned unchanged rather than
    guessed at.
    """
    if value is None or isinstance(value, bool):
        return value
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return value
    if not math.isfinite(serial):
        return value
    try:
        moment = _EXCEL_EPOCH + timedelta(days=serial)
    except OverflowError:
        return value
    if abs(serial - round(serial)) < 1e-9:
        return moment.date().isoformat()
    return moment.replace(microsecond=0).isoformat()


def _convert_date_columns(
    df: pd.DataFrame, column_types: Optional[dict[str, str]]
) -> pd.DataFrame:
    """Rewrites serial-date columns as ISO-8601 strings before the agent sees them.

    Models doing serial-to-date arithmetic by hand is an observed source of
    fabricated dates; the conversion belongs here, deterministically.
    """
    if not column_types:
        return df
    for col, col_type in column_types.items():
        if col_type == "date" and col in df.columns:
            df[col] = df[col].map(_serial_to_iso)
    return df


def _drift_report(
    live_headers: list[str],
    sheet_info: Optional[dict[str, Any]],
    live_row_count: Optional[int] = None,
) -> Optional[dict[str, Any]]:
    """Compares the freshly read header row and row count against the
    scan-time fingerprint.

    graph.json stores header_row from scan time; a row inserted above the
    header in OneDrive shifts the used range and the cached index then points
    at a data row. The mismatch is only loud when it produces garbage column
    names — this makes it loud always.

    Header drift alone misses the far more common edit on a multi-region
    sheet: rows appended or removed *inside* a table body, which never
    touches the header at all but leaves every cached region body span (and
    any label banner below it) one edit away from wrong. live_row_count —
    already read for the actual query, so this costs no extra Graph call —
    catches that case by comparing against the row count recorded at scan
    time.
    """
    sheet_info = sheet_info or {}
    stored_headers = sheet_info.get("columns")
    header_changed = bool(stored_headers) and live_headers != stored_headers

    stored_row_count = sheet_info.get("row_count")
    row_count_changed = (
        live_row_count is not None
        and stored_row_count is not None
        and live_row_count != stored_row_count
    )

    if not header_changed and not row_count_changed:
        return None

    report: dict[str, Any] = {"structure_drift": True}
    if header_changed:
        report["stored_header"] = stored_headers
        report["live_header"] = live_headers
    if row_count_changed:
        report["stored_row_count"] = stored_row_count
        report["live_row_count"] = live_row_count
        if sheet_info.get("regions"):
            report["region_drift_risk"] = (
                "This sheet's used range now has a different row count than "
                "at scan time, but its regions (table bodies, headers, "
                "labels) are cached from that same scan. Rows were likely "
                "inserted or removed inside a table body without moving the "
                "header, so region body spans below the edit may now be "
                "wrong. Treat region boundaries here as unconfirmed until "
                "scan_workspace refreshes them."
            )
    report["recommendation"] = (
        "The sheet's structure no longer matches what was captured at scan "
        "time — the sheet was probably edited. Results below use the live "
        "header, but run scan_workspace to refresh the index."
    )
    return report


async def _fetch_sheet_data(
    item_id: str,
    sheet_name: str,
    sheet_info: Optional[dict[str, Any]] = None,
) -> tuple[pd.DataFrame, Optional[dict[str, Any]]]:
    """Fetches live data and returns (frame, drift_report). Never reads from cache.

    sheet_info is the sheet's graph entry: header_row positions the header,
    column_types drives serial-date conversion, columns is the drift
    fingerprint.
    """
    sheet_info = sheet_info or {}
    header_row = sheet_info.get("header_row", 1)
    data = await get_used_range(item_id, sheet_name)
    values = data.get("values", [])
    if not values or len(values) < header_row:
        return pd.DataFrame(), None

    header_idx = header_row - 1
    # Same normalisation as scan-time discovery, so the frame's column names
    # always match what get_workspace_graph told the agent.
    headers = columns_from_row(values[header_idx])
    drift = _drift_report(headers, sheet_info, live_row_count=len(values))

    data_rows = values[header_idx + 1:]
    padded: list[list[Any]] = []
    for row in data_rows:
        if len(row) < len(headers):
            row = row + [None] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[: len(headers)]
        padded.append(row)

    df = _convert_date_columns(
        pd.DataFrame(padded, columns=headers), sheet_info.get("column_types")
    )
    return df, drift


def _date_compare(df: pd.DataFrame, col: str, bound: str, op: str) -> pd.DataFrame:
    """Compares an ISO-8601 date column against an ISO literal.

    Date columns arrive from _convert_date_columns as ISO strings, whose
    lexicographic order is chronological order. Rows that are not ISO-shaped
    (blanks, stray text) never match a date comparison.
    """
    series = df[col].astype(str).str.strip()
    is_date = series.str.match(_ISO_DATE_RE.pattern, na=False)
    if op == ">":
        return df[is_date & (series > bound)]
    if op == ">=":
        return df[is_date & (series >= bound)]
    if op == "<":
        return df[is_date & (series < bound)]
    return df[is_date & (series <= bound)]


def _numeric_compare(df: pd.DataFrame, col: str, raw: str, op: str) -> pd.DataFrame:
    """Applies a numeric or ISO-date comparison, raising rather than silently skipping.

    A malformed bound such as `">abc"` previously fell into `except: pass`, so the
    condition was dropped and the caller received the *unfiltered* sheet while
    believing it was filtered — a silently wrong answer, which is worse than an error.
    """
    raw_str = str(raw).strip()
    if _ISO_DATE_RE.match(raw_str):
        return _date_compare(df, col, raw_str, op)
    try:
        bound = float(raw)
    except (TypeError, ValueError):
        raise ValueError(
            f"Condition on column '{col}' uses '{op}' but '{raw}' is neither a "
            f"number nor an ISO date. Use e.g. {{'{col}': '{op}100'}} or "
            f"{{'{col}': '{op}2026-01-01'}}."
        )
    series = pd.to_numeric(df[col], errors="coerce")
    if op == ">":
        return df[series > bound]
    if op == ">=":
        return df[series >= bound]
    if op == "<":
        return df[series < bound]
    return df[series <= bound]


def _exact_string_match(
    series: pd.Series, val: Any, exact_case: bool
) -> pd.Series:
    """Equality mask for string comparison.

    Excel cells routinely carry trailing whitespace and inconsistent casing, so
    the default comparison strips and casefolds both sides — {"Status":
    "closed"} matches "Closed " rather than confidently returning zero rows.
    exact_case=True restores byte-for-byte matching.
    """
    if exact_case:
        return series.astype(str) == str(val)
    return series.astype(str).str.strip().str.casefold() == str(val).strip().casefold()


def _equality_mask(series: pd.Series, val: Any, exact_case: bool) -> pd.Series:
    """Mask for one equality comparison: numeric when both sides are numeric,
    normalised-string otherwise."""
    try:
        numeric_val = float(val)
        numeric_series = pd.to_numeric(series, errors="coerce")
        if not numeric_series.isna().all():
            return numeric_series == numeric_val
    except (ValueError, TypeError):
        pass
    return _exact_string_match(series, val, exact_case)


def _null_mask(series: pd.Series) -> pd.Series:
    """Excel blanks arrive as None or empty strings — both count as null."""
    return series.isna() | (series.astype(str).str.strip() == "")


_STRUCTURED_OPS = ("in", "between", "is_null", "contains", "=", ">", ">=", "<", "<=")


def _apply_structured_condition(
    df: pd.DataFrame, col: str, spec: dict[str, Any], exact_case: bool
) -> pd.DataFrame:
    """Applies one dict-form condition; multiple operators AND together, so
    {"Batch Date": {">=": "2026-01-01", "<": "2026-04-01"}} is a range."""
    if not spec:
        raise ValueError(
            f"Condition on column '{col}' is an empty object. "
            f"Use operators: {', '.join(_STRUCTURED_OPS)}."
        )
    for op, arg in spec.items():
        if op == "in":
            if not isinstance(arg, (list, tuple)) or not arg:
                raise ValueError(
                    f"'in' on column '{col}' needs a non-empty list, got {arg!r}."
                )
            mask = _equality_mask(df[col], arg[0], exact_case)
            for candidate in arg[1:]:
                mask |= _equality_mask(df[col], candidate, exact_case)
            df = df[mask]
        elif op == "between":
            if not isinstance(arg, (list, tuple)) or len(arg) != 2:
                raise ValueError(
                    f"'between' on column '{col}' needs [low, high], got {arg!r}."
                )
            df = _numeric_compare(df, col, str(arg[0]), ">=")
            df = _numeric_compare(df, col, str(arg[1]), "<=")
        elif op == "is_null":
            if not isinstance(arg, bool):
                raise ValueError(
                    f"'is_null' on column '{col}' needs true or false, got {arg!r}."
                )
            mask = _null_mask(df[col])
            df = df[mask] if arg else df[~mask]
        elif op == "contains":
            df = df[
                df[col].astype(str).str.contains(
                    str(arg), case=False, na=False, regex=False
                )
            ]
        elif op == "=":
            df = df[_equality_mask(df[col], arg, exact_case)]
        elif op in (">", ">=", "<", "<="):
            df = _numeric_compare(df, col, str(arg), op)
        else:
            raise ValueError(
                f"Unknown operator '{op}' on column '{col}'. "
                f"Use: {', '.join(_STRUCTURED_OPS)}."
            )
    return df


def _apply_conditions(
    df: pd.DataFrame, conditions: dict[str, Any], exact_case: bool = False
) -> pd.DataFrame:
    if not isinstance(conditions, dict):
        raise ValueError("conditions must be an object mapping column names to values.")

    unknown = [c for c in conditions if c not in df.columns]
    if unknown:
        raise ValueError(
            f"Condition column(s) {unknown} not found in sheet. "
            f"Available columns: {list(df.columns)}. "
            f"Call get_workspace_graph or inspect_file to see the real column names."
        )

    for col, val in conditions.items():
        if isinstance(val, dict):
            df = _apply_structured_condition(df, col, val, exact_case)
            continue

        val_str = str(val)

        if val_str.startswith("~"):
            # regex=False is essential: the substring comes straight from the
            # model/user, so regex interpretation allowed both accidental
            # mismatches (a literal "." matching any char) and catastrophic
            # backtracking that would hang the server on a crafted pattern.
            df = df[
                df[col].astype(str).str.contains(
                    val_str[1:], case=False, na=False, regex=False
                )
            ]
        elif val_str.startswith(">="):
            df = _numeric_compare(df, col, val_str[2:], ">=")
        elif val_str.startswith("<="):
            df = _numeric_compare(df, col, val_str[2:], "<=")
        elif val_str.startswith(">"):
            df = _numeric_compare(df, col, val_str[1:], ">")
        elif val_str.startswith("<"):
            df = _numeric_compare(df, col, val_str[1:], "<")
        else:
            df = df[_equality_mask(df[col], val, exact_case)]

    # Filtering yields views; downstream code assigns to columns (to_numeric
    # coercion), which raises SettingWithCopyWarning and may not propagate.
    return df.copy()


def _zero_match_diagnostics(
    df: pd.DataFrame, conditions: dict[str, Any], exact_case: bool = False
) -> dict[str, Any]:
    """Explains a zero-row result so the agent can self-correct.

    'rows: []' with no signal reads as "no such orders exist" and gets reported
    with full confidence. Showing what each condition matched alone, plus the
    values actually present in the offending column, turns that into a retry.
    """
    per_condition: list[dict[str, Any]] = []
    for col, val in conditions.items():
        matched_alone = len(_apply_conditions(df, {col: val}, exact_case))
        entry: dict[str, Any] = {
            "column": col,
            "condition": val,
            "rows_matching_this_condition_alone": matched_alone,
        }
        if matched_alone == 0:
            distinct = (
                df[col].dropna().astype(str).str.strip().drop_duplicates().head(20)
            )
            entry["distinct_values_present"] = distinct.tolist()
        per_condition.append(entry)
    return {
        "note": (
            "No rows matched the combined conditions. Conditions that matched "
            "zero rows on their own include a sample of the values actually "
            "present in that column — check for a near-miss and retry."
        ),
        "conditions": per_condition,
    }


def _get_file_info(folder_path: str, file_name: str) -> tuple[dict[str, Any], dict[str, Any]]:
    folder_path = folder_path.rstrip("/")
    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder_path, {})
    if not workspace:
        raise ValueError(
            f"Workspace '{folder_path}' not found. Run scan_workspace first."
        )
    require_current_schema(workspace, folder_path)
    file_info = workspace.get("files", {}).get(file_name)
    if not file_info:
        available = list(workspace.get("files", {}).keys())
        raise ValueError(
            f"File '{file_name}' not found in workspace '{folder_path}'. "
            f"Available: {available}"
        )
    return workspace, file_info


def _resolve_sheet(
    file_info: dict[str, Any], sheet_name: str, file_name: str
) -> dict[str, Any]:
    """Returns the sheet's graph entry, failing loudly if the sheet is unknown.

    An unknown sheet used to fall through to a default header_row of 1 and then
    produce a confusing Graph 404 instead of naming the valid sheets.
    """
    sheets = file_info.get("sheets", {})
    if sheet_name not in sheets:
        raise ValueError(
            f"Sheet '{sheet_name}' not found in '{file_name}'. "
            f"Available sheets: {list(sheets.keys())}"
        )
    return sheets[sheet_name]


def _region_summary(index: int, region: dict[str, Any]) -> dict[str, Any]:
    summary = {
        "index": index,
        "body": region.get("body"),
        "header_row": region.get("header_row"),
    }
    if region.get("label"):
        summary["label"] = region["label"]
    return summary


def _region_span_text(index: int, region: dict[str, Any]) -> str:
    """'0: NAPHTHALENE, rows 7:28' when the region has a label scanned from a
    section banner, else the bare '0: 7:28' — the difference between the
    model guessing which table answers a question and the model knowing."""
    label = region.get("label")
    body = region.get("body")
    return f"{index}: {label}, rows {body}" if label else f"{index}: {body}"


_SPAN_TEXT_RE = re.compile(r"^\s*\d+\s*:\s*\d+\s*$")


def _match_region_by_label(
    regions: list[dict[str, Any]], query: str
) -> tuple[Optional[tuple[int, dict[str, Any]]], list[dict[str, Any]]]:
    """Fuzzy-matches `query` against each region's scanned label.

    Returns (match, ambiguous_candidates). `match` is set only when exactly
    one region's label plausibly means `query` — an exact case/whitespace-
    insensitive match wins outright; otherwise the same fuzzy logic already
    used for sheet-name drift (containment, small typos, via
    fuzzy_name_candidates) decides. More than one plausible region is
    ambiguity to report, not a guess to make silently, so it comes back in
    `ambiguous_candidates` instead of being resolved.
    """
    labelled = [(i, r) for i, r in enumerate(regions) if r.get("label")]
    if not labelled:
        return None, []
    query_n = normalise_name(query)
    exact = [(i, r) for i, r in labelled if normalise_name(r["label"]) == query_n]
    if len(exact) == 1:
        return exact[0], []
    if len(exact) > 1:
        return None, [r for _, r in exact]
    candidates = set(fuzzy_name_candidates(query, [r["label"] for _, r in labelled]))
    fuzzy_matches = [(i, r) for i, r in labelled if r["label"] in candidates]
    if len(fuzzy_matches) == 1:
        return fuzzy_matches[0], []
    if len(fuzzy_matches) > 1:
        return None, [r for _, r in fuzzy_matches]
    return None, []


def _resolve_region(
    sheet_info: dict[str, Any], region: Any, file_name: str, sheet_name: str
) -> Optional[tuple[int, dict[str, Any]]]:
    """Resolves `region` to (index, region_entry) in sheet_info['regions'].

    `region` is an integer index, a 'first:last' absolute-row span matching a
    region's body, or a label string fuzzy-matched against each region's
    scanned section-banner label ("naphthalene" reaches "NAPHTHALENE"). None
    (the default) resolves to None — no scoping. An index out of range, a
    span matching nothing, a label matching nothing, or a label matching more
    than one region is a caller mistake, not a request to fall back to the
    whole sheet, so all of these raise with the available regions listed
    rather than silently ignoring `region` or guessing among candidates.
    """
    if region is None:
        return None
    regions = sheet_info.get("regions") or []
    if isinstance(region, bool):
        raise ValueError(
            f"region must be an integer index, a 'first:last' span string, "
            f"or a region label, got {region!r}."
        )
    if isinstance(region, int):
        if 0 <= region < len(regions):
            return region, regions[region]
        raise ValueError(
            f"region index {region} is out of range for '{file_name}'/"
            f"'{sheet_name}': it has {len(regions)} region(s): "
            f"{[_region_summary(i, r) for i, r in enumerate(regions)]}"
        )
    if isinstance(region, str):
        text = region.strip()
        if _SPAN_TEXT_RE.match(text):
            wanted = parse_span(text)
            for i, r in enumerate(regions):
                body = r.get("body")
                if body and parse_span(body) == wanted:
                    return i, r
            raise ValueError(
                f"region '{region}' matches no region body on '{file_name}'/"
                f"'{sheet_name}'. Available regions: "
                f"{[_region_summary(i, r) for i, r in enumerate(regions)]}"
            )
        match, ambiguous = _match_region_by_label(regions, text)
        if match:
            return match
        if ambiguous:
            raise ValueError(
                f"region label '{region}' matches more than one region on "
                f"'{file_name}'/'{sheet_name}': "
                f"{[r.get('label') for r in ambiguous]}. Use the region "
                f"index instead. Available regions: "
                f"{[_region_summary(i, r) for i, r in enumerate(regions)]}"
            )
        raise ValueError(
            f"region '{region}' matches no region body or label on "
            f"'{file_name}'/'{sheet_name}'. Available regions: "
            f"{[_region_summary(i, r) for i, r in enumerate(regions)]}"
        )
    raise ValueError(
        f"region must be an integer index, a 'first:last' span string, or a "
        f"region label, got {region!r}."
    )


def _scope_to_region(
    df: pd.DataFrame,
    sheet_info: dict[str, Any],
    region: Any,
    file_name: str,
    sheet_name: str,
    *,
    allow_full_sheet_fallback: bool = False,
) -> tuple[pd.DataFrame, Optional[dict[str, Any]], Optional[str]]:
    """Slices `df` to one table region when asked.

    Default behavior is strict: multi-region sheets require an explicit region,
    otherwise a ValueError is raised so the caller cannot silently flatten
    multiple subtables together. When `allow_full_sheet_fallback=True`, the
    caller intentionally opts into the warning-only whole-sheet read path.
    """
    regions = sheet_info.get("regions") or []
    resolved = _resolve_region(sheet_info, region, file_name, sheet_name)
    if resolved is None:
        if len(regions) > 1:
            spans = "; ".join(
                _region_span_text(i, r) for i, r in enumerate(regions)
            )
            if not allow_full_sheet_fallback:
                raise ValueError(
                    f"'{sheet_name}' in '{file_name}' holds {len(regions)} table "
                    f"regions ({spans}) and no `region` was given. Pass "
                    f"region=<index> or region='<body span>' to scope to one table; "
                    f"otherwise set allow_full_sheet_fallback=True for an explicit "
                    f"whole-sheet fallback."
                )
            warning = (
                f"'{sheet_name}' in '{file_name}' holds {len(regions)} table "
                f"regions ({spans}) and a whole-sheet fallback was explicitly "
                f"requested. This result spans all of them. Pass region=<index> "
                f"or region='<body span>' to scope to one table."
            )
            return df, None, warning
        return df, None, None

    index, entry = resolved
    used_range_address = sheet_info.get("used_range_address")
    if not used_range_address:
        raise ValueError(
            f"'{sheet_name}' in '{file_name}' has no used_range_address on "
            f"record — run scan_workspace to rebuild the structure graph."
        )
    header_row = sheet_info.get("header_row", 1)
    start, stop = region_to_df_slice(used_range_address, header_row, entry["body"])
    start = max(start, 0)
    stop = max(min(stop, len(df)), start)
    scoped = df.iloc[start:stop].copy()
    region_used: dict[str, Any] = {
        "index": index,
        "body": entry.get("body"),
        "row_count": len(scoped),
    }
    if entry.get("label"):
        region_used["label"] = entry["label"]
    return scoped, region_used, None


def _scope_to_region_for_derive(
    df: pd.DataFrame,
    sheet_info: dict[str, Any],
    region: Any,
    file_name: str,
    sheet_name: str,
    *,
    allow_full_sheet_fallback: bool = False,
) -> tuple[pd.DataFrame, Optional[dict[str, Any]], Optional[str]]:
    """Like `_scope_to_region`, but a resolved region also pulls in its
    `pre_body` rows — an opening balance sitting between the header and the
    body, which the region's own formulas never totalled and body-only
    scoping therefore excludes by construction.

    Only `derive` calls this. `filter_sheet` and `aggregate` keep calling
    `_scope_to_region` unchanged: a raw row dump or a sum of the body has no
    business including a row the sheet itself never summed, and widening
    their frame silently would make a total that used to exclude that row
    start including it without anyone asking.
    """
    scoped, region_used, region_warning = _scope_to_region(
        df,
        sheet_info,
        region,
        file_name,
        sheet_name,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )
    if region_used is None:
        return scoped, region_used, region_warning

    regions = sheet_info.get("regions") or []
    index = region_used.get("index")
    entry = regions[index] if isinstance(index, int) and 0 <= index < len(regions) else None
    pre_body = entry.get("pre_body") if entry else None
    if not pre_body:
        return scoped, region_used, region_warning

    used_range_address = sheet_info.get("used_range_address")
    header_row = sheet_info.get("header_row", 1)
    start, stop = region_to_df_slice(used_range_address, header_row, pre_body)
    start = max(start, 0)
    stop = max(min(stop, len(df)), start)
    pre_body_df = df.iloc[start:stop]
    if pre_body_df.empty:
        return scoped, region_used, region_warning

    widened = pd.concat([pre_body_df, scoped], ignore_index=True)
    region_used = dict(region_used)
    region_used["pre_body"] = pre_body
    region_used["row_count"] = len(widened)
    return widened, region_used, region_warning


def _clamp_limit(limit: Optional[int]) -> int:
    """Normalises a caller-supplied row limit into [1, MAX_ROWS_RETURNED]."""
    if limit is None:
        return MAX_ROWS_RETURNED
    try:
        limit = int(limit)
    except (TypeError, ValueError):
        raise ValueError(f"limit must be an integer, got {limit!r}.")
    if limit <= 0:
        # df.head(-5) drops rows from the end, which silently returns the wrong
        # set rather than erroring.
        raise ValueError(f"limit must be a positive integer, got {limit}.")
    return min(limit, MAX_ROWS_RETURNED)


# Two routed sheets whose combined scores sit within this delta are
# indistinguishable to the router; the agent should be told, not left to
# treat an arbitrary pick as a confident one.
ROUTING_TIE_DELTA = 0.02


def _routing_summary(results: list[dict]) -> dict[str, Any]:
    scores = [r.get("score", 0.0) for r in results]
    near_ties = [
        {"file": r["file"], "sheet": r["sheet"], "score": round(r.get("score", 0.0), 4)}
        for r in results
        if scores and scores[0] - r.get("score", 0.0) < ROUTING_TIE_DELTA
    ]
    ambiguous = len(near_ties) > 1
    summary: dict[str, Any] = {"routing_ambiguous": ambiguous}
    if ambiguous:
        summary["near_ties"] = near_ties
        summary["note"] = (
            "The top-ranked sheets scored within a tie margin of each other — "
            "the routing choice between them is effectively arbitrary. "
            "Confirm with inspect_file or ask the user which file is meant "
            "before trusting an answer from just one of them."
        )
    return summary


async def execute_query(
    question: str,
    folder_path: str,
    n_results: int = 5,
    min_score: Optional[float] = None,
) -> dict:
    folder_path = folder_path.rstrip("/")
    if not isinstance(n_results, int) or n_results < 1:
        raise ValueError(f"n_results must be a positive integer, got {n_results!r}.")
    results = search(folder_path, question, n_results=min(n_results, 20))
    if min_score is not None:
        results = [r for r in results if r.get("score", 0.0) >= min_score]
    if not results:
        return {
            "results": [],
            "files": [],
            "sheets": [],
            "row_count": 0,
            "message": (
                "No relevant sheets found for this workspace. "
                "Run scan_workspace first, check the folder_path, or lower "
                "min_score."
            ),
        }

    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder_path, {})

    async def fetch_and_format(meta: dict) -> dict:
        file_name = meta["file"]
        sheet_name = meta["sheet"]
        file_info = workspace.get("files", {}).get(file_name, {})
        item_id = file_info.get("item_id", "")
        if not item_id:
            return {
                "file": file_name,
                "sheet": sheet_name,
                "score": meta.get("score", 0),
                "rows": [],
                "row_count": 0,
                "error": "File is in the search index but missing from the "
                         "structure graph. Run scan_workspace.",
            }
        sheet_info = file_info.get("sheets", {}).get(sheet_name, {})
        df, drift = await _fetch_sheet_data(item_id, sheet_name, sheet_info)
        result = {
            "file": file_name,
            "sheet": sheet_name,
            "score": meta.get("score", 0),
            "rows": _records(df.head(QUERY_ROWS_PER_SHEET)),
            "row_count": len(df),
            "truncated": len(df) > QUERY_ROWS_PER_SHEET,
        }
        if drift:
            result["structure_drift"] = drift
        return result

    raw = await asyncio.gather(
        *[fetch_and_format(m) for m in results], return_exceptions=True
    )

    sheet_results: list[dict] = []
    for meta, item in zip(results, raw):
        if isinstance(item, Exception):
            # One unreadable sheet should not fail the whole exploratory query.
            sheet_results.append(
                {
                    "file": meta["file"],
                    "sheet": meta["sheet"],
                    "score": meta.get("score", 0),
                    "rows": [],
                    "row_count": 0,
                    "error": str(item),
                }
            )
        else:
            sheet_results.append(item)

    return {
        "results": sheet_results,
        "routing": _routing_summary(results),
        "files": [r["file"] for r in sheet_results],
        "sheets": [r["sheet"] for r in sheet_results],
        "row_count": sum(r["row_count"] for r in sheet_results),
    }


async def execute_inspect_file(file_name: str, folder_path: str) -> dict:
    folder_path = folder_path.rstrip("/")
    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder_path, {})
    if not workspace:
        return {"error": f"Workspace '{folder_path}' not found. Run scan_workspace first."}
    file_info = workspace.get("files", {}).get(file_name)
    if not file_info:
        available = list(workspace.get("files", {}).keys())
        return {"error": f"File '{file_name}' not found.", "available": available}

    return {
        "file": file_name,
        "item_id": file_info.get("item_id", ""),
        "last_scanned": file_info.get("last_scanned", ""),
        "row_counts_are": "as_of_last_scan",
        "sheets": {
            name: {
                "header_row": s.get("header_row", 1),
                "header_source": s.get("header_source", "heuristic"),
                "columns": s.get("columns", []),
                "approx_row_count": s.get("approx_row_count"),
                "column_types": s.get("column_types", {}),
                # Absolute sheet rows, unlike header_row. A sheet with more
                # than one table region is one where an unqualified aggregate
                # spans blocks that were never meant to be added together.
                "regions": s.get("regions", []),
                "unclaimed_rows": s.get("unclaimed_rows", []),
                "layout_confidence": s.get("layout_confidence", "unconfirmed"),
            }
            for name, s in file_info.get("sheets", {}).items()
        },
    }


async def execute_filter_sheet(
    file_name: str,
    sheet_name: str,
    conditions: dict[str, Any],
    sort_by: Optional[str],
    limit: Optional[int],
    folder_path: str,
    exact_case: bool = False,
    region: Any = None,
    *,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder_path = folder_path.rstrip("/")
    effective_limit = _clamp_limit(limit)
    _, file_info = _get_file_info(folder_path, file_name)
    sheet_info = _resolve_sheet(file_info, sheet_name, file_name)
    full_df, drift = await _fetch_sheet_data(
        file_info["item_id"], sheet_name, sheet_info
    )
    full_df, region_used, region_warning = _scope_to_region(
        full_df,
        sheet_info,
        region,
        file_name,
        sheet_name,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )

    df = full_df
    if conditions:
        df = _apply_conditions(df, conditions, exact_case)

    if sort_by:
        if sort_by not in df.columns:
            raise ValueError(
                f"sort_by column '{sort_by}' not found. Available: {list(df.columns)}"
            )
        df = df.sort_values(by=sort_by)

    total_matched = len(df)
    result_df = df.head(effective_limit)
    result = {
        "rows": _records(result_df),
        "file": file_name,
        "sheet": sheet_name,
        "row_count": len(result_df),
        "total_matched": total_matched,
        "truncated": total_matched > len(result_df),
    }
    if total_matched == 0 and conditions and not full_df.empty:
        result["zero_match_diagnostics"] = _zero_match_diagnostics(
            full_df, conditions, exact_case
        )
    if drift:
        result["structure_drift"] = drift
    if region_used:
        result["region_used"] = region_used
    if region_warning:
        result["region_warning"] = region_warning
    return result


async def execute_sheet_layout(
    file_name: str,
    sheet_name: str,
    folder_path: str,
) -> dict:
    """Returns a workbook-layout summary for one sheet, including region spans."""
    folder_path = folder_path.rstrip("/")
    _, file_info = _get_file_info(folder_path, file_name)
    sheet_info = _resolve_sheet(file_info, sheet_name, file_name)
    df, drift = await _fetch_sheet_data(file_info["item_id"], sheet_name, sheet_info)

    region_entries: list[dict[str, Any]] = []
    used_range_address = sheet_info.get("used_range_address")
    header_row = sheet_info.get("header_row", 1)
    for index, entry in enumerate(sheet_info.get("regions") or []):
        body = entry.get("body")
        row_count = 0
        if body and used_range_address:
            try:
                start, stop = region_to_df_slice(used_range_address, header_row, body)
                start = max(start, 0)
                stop = max(min(stop, len(df)), start)
                row_count = stop - start
            except (TypeError, ValueError):
                row_count = 0
        region_info: dict[str, Any] = {
            "index": index,
            "body": body,
            "header_row": entry.get("header_row"),
            "formula_header_row": entry.get("formula_header_row"),
            "range": entry.get("range"),
            "source": entry.get("source"),
            "confidence": entry.get("confidence"),
            "row_count": row_count,
        }
        if entry.get("pre_body"):
            region_info["pre_body"] = entry["pre_body"]
        if entry.get("label"):
            region_info["label"] = entry["label"]
        region_entries.append(region_info)

    result = {
        "file": file_name,
        "sheet": sheet_name,
        "header_row": header_row,
        "columns": sheet_info.get("columns", list(df.columns)),
        "column_types": sheet_info.get("column_types", {}),
        "used_range_address": used_range_address,
        "row_count": len(df),
        "region_count": len(region_entries),
        "regions": region_entries,
        "unclaimed_rows": sheet_info.get("unclaimed_rows", []),
        "layout_confidence": sheet_info.get("layout_confidence", "unconfirmed"),
    }
    if drift:
        result["structure_drift"] = drift
    return result


async def execute_fetch_sheet_rows(
    file_name: str,
    sheet_name: str,
    folder_path: str,
    limit: Optional[int] = 500,
    region: Any = None,
    *,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    """Returns the live rows for a sheet or a scoped region without the hard cap
    used by general row-returning queries.

    The default limit keeps the response readable for complex workbooks while the
    caller can pass limit=None to fetch the full region/sheet. This tool is a
    factual raw-data dump, so it does not infer or guess structure; it only
    fetches what the workbook actually contains and labels the region choice.
    """
    folder_path = folder_path.rstrip("/")
    _, file_info = _get_file_info(folder_path, file_name)
    sheet_info = _resolve_sheet(file_info, sheet_name, file_name)
    full_df, drift = await _fetch_sheet_data(
        file_info["item_id"], sheet_name, sheet_info
    )
    scoped_df, region_used, region_warning = _scope_to_region(
        full_df,
        sheet_info,
        region,
        file_name,
        sheet_name,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )

    if limit is not None:
        try:
            limit = int(limit)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"limit must be an integer or None, got {limit!r}.") from exc
        if limit <= 0:
            raise ValueError(f"limit must be a positive integer, got {limit}.")
        result_df = scoped_df.head(limit)
        truncated = len(scoped_df) > len(result_df)
    else:
        result_df = scoped_df.copy()
        truncated = False

    result = {
        "rows": _records(result_df),
        "file": file_name,
        "sheet": sheet_name,
        "row_count": len(result_df),
        "total_matched": len(scoped_df),
        "truncated": truncated,
    }
    if drift:
        result["structure_drift"] = drift
    if region_used:
        result["region_used"] = region_used
    if region_warning:
        result["region_warning"] = region_warning
    return result


async def execute_fetch_region_rows(
    file_name: str,
    sheet_name: str,
    folder_path: str,
    region: Any,
    limit: Optional[int] = None,
) -> dict:
    """Fetches one exact region, never falling back to the whole sheet."""
    if region is None:
        raise ValueError(
            "region is required for fetch_region_rows; pass the region index or body span "
            "such as 0 or '7:28'."
        )
    result = await execute_fetch_sheet_rows(
        file_name, sheet_name, folder_path, limit=limit, region=region
    )
    if "region_used" not in result:
        raise ValueError(
            f"No region resolved for '{file_name}'/'{sheet_name}' with region={region!r}. "
            "Pass an exact region index or body span from the sheet layout."
        )
    return result


async def execute_aggregate(
    file_name: str,
    sheet_name: str,
    group_by: Any,
    value_col: str,
    operation: str,
    conditions: Optional[dict[str, Any]],
    folder_path: str,
    having: Optional[dict[str, Any]] = None,
    region: Any = None,
    *,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder_path = folder_path.rstrip("/")
    _, file_info = _get_file_info(folder_path, file_name)
    sheet_info = _resolve_sheet(file_info, sheet_name, file_name)
    full_df, drift = await _fetch_sheet_data(
        file_info["item_id"], sheet_name, sheet_info
    )
    full_df, region_used, region_warning = _scope_to_region(
        full_df,
        sheet_info,
        region,
        file_name,
        sheet_name,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )

    df = full_df
    if conditions:
        df = _apply_conditions(df, conditions)

    ops_allowed = {"sum", "mean", "count", "min", "max"}
    if operation not in ops_allowed:
        raise ValueError(
            f"Unknown operation '{operation}'. Use: {', '.join(sorted(ops_allowed))}."
        )
    group_cols = group_by if isinstance(group_by, list) else [group_by]
    if not group_cols or not all(isinstance(g, str) and g for g in group_cols):
        raise ValueError(
            f"group_by must be a column name or a list of them, got {group_by!r}."
        )
    if value_col not in df.columns:
        raise ValueError(f"Column '{value_col}' not found. Available: {list(df.columns)}")
    missing_groups = [g for g in group_cols if g not in df.columns]
    if missing_groups:
        raise ValueError(
            f"group_by column(s) {missing_groups} not found. "
            f"Available: {list(df.columns)}"
        )

    if operation != "count":
        df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

    result = getattr(df.groupby(group_cols)[value_col], operation)().reset_index()
    if having:
        # The having clause filters aggregated rows with the same grammar as
        # conditions — unknown columns still raise, naming what exists.
        result = _apply_conditions(result, having)
    response = {
        "rows": _records(result.head(MAX_ROWS_RETURNED)),
        "file": file_name,
        "sheet": sheet_name,
        "row_count": min(len(result), MAX_ROWS_RETURNED),
        "truncated": len(result) > MAX_ROWS_RETURNED,
    }
    if df.empty and conditions and not full_df.empty:
        response["zero_match_diagnostics"] = _zero_match_diagnostics(
            full_df, conditions
        )
    if drift:
        response["structure_drift"] = drift
    if region_used:
        response["region_used"] = region_used
    if region_warning:
        response["region_warning"] = region_warning
    return response


async def execute_cross_file_aggregate(
    folder_path: str,
    sheet_name: str,
    value_col: str,
    operation: str,
    conditions: Optional[dict[str, Any]] = None,
) -> dict:
    folder_path = folder_path.rstrip("/")
    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder_path, {})
    if not workspace:
        raise ValueError(
            f"Workspace '{folder_path}' not found. Run scan_workspace first."
        )

    # Files without the exact sheet used to be silently excluded from the
    # total — no warning, no listing — so a workspace where some files call the
    # sheet 'Sales' and others 'Sales 2024' produced a confidently wrong total.
    # They are still excluded (guessing would be worse), but now visibly.
    matching: list[tuple[str, dict]] = []
    unmatched_files: list[dict[str, Any]] = []
    for fn, fi in workspace.get("files", {}).items():
        sheets = fi.get("sheets", {})
        if sheet_name in sheets:
            matching.append((fn, fi))
        else:
            entry: dict[str, Any] = {"file": fn, "sheets": list(sheets.keys())}
            candidates = fuzzy_name_candidates(sheet_name, sheets.keys())
            if candidates:
                entry["did_you_mean"] = candidates
            unmatched_files.append(entry)

    if not matching:
        all_candidates = sorted(
            {c for u in unmatched_files for c in u.get("did_you_mean", [])}
        )
        hint = (
            f" Close matches in this workspace: {all_candidates}."
            if all_candidates
            else ""
        )
        raise ValueError(
            f"No files found with sheet '{sheet_name}' in workspace "
            f"'{folder_path}'.{hint}"
        )

    ops_allowed = {"sum", "mean", "count", "min", "max"}
    if operation not in ops_allowed:
        raise ValueError(
            f"Unknown operation '{operation}'. Use: {', '.join(sorted(ops_allowed))}."
        )

    async def fetch_file(file_name: str, file_info: dict):
        try:
            sheet_info = file_info["sheets"][sheet_name]
            df, drift = await _fetch_sheet_data(
                file_info["item_id"], sheet_name, sheet_info
            )
            if conditions:
                df = _apply_conditions(df, conditions)
            return file_name, df, drift, None
        except Exception as exc:
            return file_name, None, None, exc

    # Results are folded incrementally as each file completes: one frame is
    # held at a time instead of concatenating all of them. For mean, (sum,
    # count) pairs accumulate so the result stays row-weighted — averaging
    # per-file means would weight a 3-row file the same as a 30,000-row one.
    # Concurrency is bounded by the Graph client's shared semaphore.
    per_file: dict[str, Any] = {}
    failures: list[dict[str, str]] = []
    missing_col: list[str] = []
    drifted: list[dict[str, Any]] = []
    total_sum = 0.0
    total_count = 0
    running_min: Optional[float] = None
    running_max: Optional[float] = None
    row_count = 0

    for future in asyncio.as_completed(
        [fetch_file(fn, fi) for fn, fi in matching]
    ):
        file_name, df, drift, exc = await future
        if exc is not None:
            failures.append({"file": file_name, "error": str(exc)})
            log(f"[Warning] Skipped {file_name}: {exc}")
            continue
        if drift:
            drifted.append({"file": file_name, **drift})
        if value_col not in df.columns:
            missing_col.append(file_name)
            continue
        row_count += len(df)

        if operation == "count":
            file_count = int(df[value_col].count())
            per_file[file_name] = float(file_count)
            total_count += file_count
            continue

        series = pd.to_numeric(df[value_col], errors="coerce")
        file_val = series.agg(operation)
        per_file[file_name] = float(file_val) if pd.notna(file_val) else None
        total_sum += float(series.sum())
        file_count = int(series.count())
        total_count += file_count
        if file_count:
            file_min, file_max = float(series.min()), float(series.max())
            running_min = (
                file_min if running_min is None else min(running_min, file_min)
            )
            running_max = (
                file_max if running_max is None else max(running_max, file_max)
            )

    if failures and len(failures) == len(matching):
        raise GraphAPIError(
            "All files failed to fetch. Cannot aggregate. "
            "Check OneDrive connectivity and try again. "
            + "; ".join(f"{f['file']}: {f['error']}" for f in failures)
        )
    if not per_file:
        raise ValueError(
            f"Column '{value_col}' not found in any file with sheet '{sheet_name}'."
        )

    if operation == "sum":
        total: Optional[float] = total_sum
    elif operation == "count":
        total = float(total_count)
    elif operation == "mean":
        total = total_sum / total_count if total_count else None
    elif operation == "min":
        total = running_min
    else:
        total = running_max

    # Completion order is nondeterministic; report in workspace order.
    file_order = {fn: i for i, (fn, _) in enumerate(matching)}
    per_file = {
        fn: per_file[fn] for fn, _ in matching if fn in per_file
    }
    failures.sort(key=lambda f: file_order[f["file"]])
    missing_col.sort(key=lambda fn: file_order[fn])
    drifted.sort(key=lambda d: file_order[d["file"]])

    warnings: list[str] = []
    if unmatched_files:
        with_hints = [u["file"] for u in unmatched_files if u.get("did_you_mean")]
        hint = (
            f" Of those, {len(with_hints)} have similarly named sheets "
            f"(see did_you_mean): {', '.join(with_hints)}."
            if with_hints
            else ""
        )
        warnings.append(
            f"{len(unmatched_files)} file(s) in this workspace have no sheet "
            f"named '{sheet_name}' and are NOT included in this total: "
            + ", ".join(u["file"] for u in unmatched_files)
            + f".{hint} The total may be incomplete — check unmatched_files."
        )
    if failures:
        warnings.append(
            f"{len(failures)} file(s) skipped due to live fetch errors: "
            + ", ".join(f["file"] for f in failures)
            + ". Total reflects available files only. Re-run to retry."
        )
    if missing_col:
        warnings.append(
            f"{len(missing_col)} file(s) had no column '{value_col}' and were "
            f"excluded: {', '.join(missing_col)}."
        )

    return {
        "total": total,
        "operation": operation,
        "column": value_col,
        "per_file": per_file,
        "files": list(per_file.keys()),
        "sheet": sheet_name,
        "row_count": row_count,
        "skipped_files": failures,
        "unmatched_files": unmatched_files,
        "drifted_files": drifted,
        "warning": " ".join(warnings) if warnings else None,
    }


# ----------------------------------------------------------------- joining


_JOIN_TYPES = {"inner", "left", "right", "outer"}


def _join_key_value(value: Any) -> Any:
    """Canonical join key: numbers by numeric value, strings normalised.

    Excel key columns mix '45', 45 and 45.0, and carry stray whitespace and
    case differences; joining on raw values silently drops matching rows.
    """
    if value is None:
        return None
    if not isinstance(value, str):
        try:
            return repr(float(value))
        except (TypeError, ValueError):
            return normalise_name(value) or None
    text = normalise_name(value)
    if not text:
        return None
    try:
        return repr(float(text))
    except ValueError:
        return text


def _suggest_join_keys(
    folder_path: str,
    workspace: dict[str, Any],
    left_file: str,
    left_sheet: str,
    right_file: str,
    right_sheet: str,
) -> tuple[str, str, dict[str, Any]]:
    """Resolves join keys from known relationships, refusing rather than guessing."""
    candidates = []
    for rel in workspace_relationships(workspace):
        # Formula-derived reference edges say "this cell reads that cell", not
        # "these columns hold the same entities". They carry confidence 1.0 and
        # would otherwise win the join outright, on a key that means nothing.
        if rel.get("kind") == "reference":
            continue
        ends = (rel.get("left", {}), rel.get("right", {}))
        for a, b in (ends, ends[::-1]):
            if (
                a.get("file") == left_file
                and a.get("sheet") == left_sheet
                and b.get("file") == right_file
                and b.get("sheet") == right_sheet
            ):
                candidates.append(
                    {
                        "left_on": a["column"],
                        "right_on": b["column"],
                        "confidence": rel.get("confidence", 0.0),
                        "source": rel.get("source", "inferred"),
                    }
                )

    declared = [c for c in candidates if c["source"] == "declared"]
    confident = [c for c in candidates if c["confidence"] >= 0.5]
    if declared:
        chosen = declared[0]
    elif len(confident) == 1:
        chosen = confident[0]
    elif len(confident) > 1:
        raise ValueError(
            f"Several plausible join keys exist between "
            f"'{left_file}'/'{left_sheet}' and '{right_file}'/'{right_sheet}': "
            f"{[(c['left_on'], c['right_on'], c['confidence']) for c in confident]}. "
            f"Pass left_on/right_on explicitly."
        )
    else:
        hint = (
            f" Low-confidence candidates: "
            f"{[(c['left_on'], c['right_on'], c['confidence']) for c in candidates]}."
            if candidates
            else ""
        )
        raise ValueError(
            f"No confident relationship is known between "
            f"'{left_file}'/'{left_sheet}' and '{right_file}'/'{right_sheet}'. "
            f"Pass left_on/right_on explicitly, or declare the relationship in "
            f"~/.excelmcp/relationships.yaml.{hint}"
        )
    return chosen["left_on"], chosen["right_on"], chosen


async def execute_join_sheets(
    folder_path: str,
    left_file: str,
    left_sheet: str,
    right_file: str,
    right_sheet: str,
    left_on: Optional[str] = None,
    right_on: Optional[str] = None,
    join_type: str = "inner",
    limit: Optional[int] = None,
) -> dict:
    folder_path = folder_path.rstrip("/")
    if join_type not in _JOIN_TYPES:
        raise ValueError(
            f"Unknown join_type '{join_type}'. Use: {', '.join(sorted(_JOIN_TYPES))}."
        )
    if (left_on is None) != (right_on is None):
        raise ValueError("Pass both left_on and right_on, or neither.")
    effective_limit = _clamp_limit(limit)

    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder_path, {})
    if not workspace:
        raise ValueError(
            f"Workspace '{folder_path}' not found. Run scan_workspace first."
        )

    _, left_info = _get_file_info(folder_path, left_file)
    left_sheet_info = _resolve_sheet(left_info, left_sheet, left_file)
    _, right_info = _get_file_info(folder_path, right_file)
    right_sheet_info = _resolve_sheet(right_info, right_sheet, right_file)

    keys_meta: dict[str, Any]
    if left_on is None:
        left_on, right_on, keys_meta = _suggest_join_keys(
            folder_path, workspace, left_file, left_sheet, right_file, right_sheet
        )
    else:
        keys_meta = {
            "left_on": left_on,
            "right_on": right_on,
            "source": "caller",
            "confidence": None,
        }

    (left_df, left_drift), (right_df, right_drift) = await asyncio.gather(
        _fetch_sheet_data(left_info["item_id"], left_sheet, left_sheet_info),
        _fetch_sheet_data(right_info["item_id"], right_sheet, right_sheet_info),
    )

    if left_on not in left_df.columns:
        raise ValueError(
            f"Join column '{left_on}' not found in '{left_file}'/'{left_sheet}'. "
            f"Available: {list(left_df.columns)}"
        )
    if right_on not in right_df.columns:
        raise ValueError(
            f"Join column '{right_on}' not found in '{right_file}'/'{right_sheet}'. "
            f"Available: {list(right_df.columns)}"
        )

    left_df = left_df.copy()
    right_df = right_df.copy()
    left_df["__excelmcp_key"] = left_df[left_on].map(_join_key_value)
    right_df["__excelmcp_key"] = right_df[right_on].map(_join_key_value)
    # Null keys never join — a blank cell matching every other blank cell is
    # a cartesian explosion, not a relationship.
    left_df = left_df[left_df["__excelmcp_key"].notna()]
    right_df = right_df[right_df["__excelmcp_key"].notna()]

    merged = left_df.merge(
        right_df,
        how=join_type,
        on="__excelmcp_key",
        suffixes=("_left", "_right"),
    ).drop(columns="__excelmcp_key")

    total_matched = len(merged)
    result_df = merged.head(effective_limit)
    result: dict[str, Any] = {
        "rows": _records(result_df),
        "keys": keys_meta,
        "join_type": join_type,
        "left": {"file": left_file, "sheet": left_sheet},
        "right": {"file": right_file, "sheet": right_sheet},
        "row_count": len(result_df),
        "total_matched": total_matched,
        "truncated": total_matched > len(result_df),
    }
    drift = [
        {"file": f, **d}
        for f, d in ((left_file, left_drift), (right_file, right_drift))
        if d
    ]
    if drift:
        result["drifted_files"] = drift
    return result


# ------------------------------------------------------------------ derive


async def execute_derive(
    folder_path: str,
    file_name: str,
    sheet_name: str,
    # group_by, quantity_col and components are all conceptually required —
    # every current caller passes them — but Python requires that anything
    # after the first defaulted positional (group_by, made optional per RULE)
    # carry a default too. Their validation below still raises when missing.
    group_by: Any = None,
    quantity_col: Optional[str] = None,
    components: Optional[list[dict[str, Any]]] = None,
    conditions: Optional[dict[str, Any]] = None,
    region: Any = None,
    *,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    """Signed sum over transaction types: the five-call stock calculation in one.

    Each component is {"conditions": {...}, "sign": +1|-1, "label": ...}; the
    result is sum(sign * groupwise_sum(quantity)) per grouping key, computed in
    pandas where it is deterministic instead of in the model where it is not.

    A component may use "base": true instead of "sign" — shorthand for sign 1,
    meant for a one-off row like an opening balance (conditions matching it,
    e.g. a blank transaction-type column, are the caller's to state; nothing
    here guesses which row that is). It still goes through the same
    conditions/sign pipeline and the same matched_no_rows flagging as any
    other component, so a mistyped opening-balance condition is reported the
    same honest way a mistyped transaction type already is.

    group_by is optional: omit it to get one net total over every matching
    row instead of one row per group.

    On a multi-region sheet, region is effectively required: omitting it
    raises (naming the available regions) rather than silently netting
    several tables together. allow_full_sheet_fallback=True opts into the
    old warn-and-continue behaviour instead, matching filter_sheet/aggregate.
    """
    folder_path = folder_path.rstrip("/")
    if not isinstance(components, list) or not components:
        raise ValueError(
            "components must be a non-empty list of "
            "{'conditions': {...}, 'sign': 1 or -1, 'label': optional}."
        )
    signs: list[int] = []
    for i, comp in enumerate(components):
        if not isinstance(comp, dict) or not isinstance(comp.get("conditions"), dict):
            raise ValueError(f"components[{i}] needs a 'conditions' object.")
        sign = comp.get("sign")
        if sign is None and comp.get("base") is True:
            sign = 1
        if sign not in (1, -1):
            raise ValueError(
                f"components[{i}] needs 'sign': 1 or -1 (or 'base': true, which "
                f"implies sign 1, for a one-off row like an opening balance)."
            )
        signs.append(sign)

    group_cols: Optional[list[str]] = None
    if group_by is not None:
        group_cols = group_by if isinstance(group_by, list) else [group_by]
        if not group_cols or not all(isinstance(g, str) and g for g in group_cols):
            raise ValueError(
                f"group_by must be a column name or a list of them, got {group_by!r}."
            )

    _, file_info = _get_file_info(folder_path, file_name)
    sheet_info = _resolve_sheet(file_info, sheet_name, file_name)
    df, drift = await _fetch_sheet_data(
        file_info["item_id"], sheet_name, sheet_info
    )
    df, region_used, region_warning = _scope_to_region_for_derive(
        df,
        sheet_info,
        region,
        file_name,
        sheet_name,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )

    if quantity_col not in df.columns:
        raise ValueError(
            f"Column '{quantity_col}' not found. Available: {list(df.columns)}"
        )
    if group_cols:
        missing_groups = [g for g in group_cols if g not in df.columns]
        if missing_groups:
            raise ValueError(
                f"group_by column(s) {missing_groups} not found. "
                f"Available: {list(df.columns)}"
            )

    if conditions:
        df = _apply_conditions(df, conditions)
    df[quantity_col] = pd.to_numeric(df[quantity_col], errors="coerce")

    net: Optional[pd.Series] = None
    net_scalar = 0.0
    breakdown = []
    for i, (comp, sign) in enumerate(zip(components, signs)):
        part = _apply_conditions(df, comp["conditions"])
        if group_cols:
            sums = part.groupby(group_cols)[quantity_col].sum() * sign
            subtotal = float(sums.sum()) if len(sums) else 0.0
            net = sums if net is None else net.add(sums, fill_value=0.0)
        else:
            subtotal = float(part[quantity_col].sum()) * sign if len(part) else 0.0
            net_scalar += subtotal
        breakdown.append(
            {
                "label": comp.get("label") or f"component_{i}",
                "sign": sign,
                "rows_matched": int(len(part)),
                "subtotal": subtotal,
                # A component matching nothing is how a typo'd transaction
                # type silently zeroes a term — say so explicitly.
                "matched_no_rows": len(part) == 0,
            }
        )

    if group_cols:
        result_frame = (net if net is not None else pd.Series(dtype=float)).reset_index()
        if len(result_frame.columns) == len(group_cols) + 1:
            result_frame.columns = [*group_cols, "net"]
    else:
        result_frame = pd.DataFrame([{"net": net_scalar}])

    response: dict[str, Any] = {
        "rows": _records(result_frame.head(MAX_ROWS_RETURNED)),
        "components": breakdown,
        "quantity_col": quantity_col,
        "file": file_name,
        "sheet": sheet_name,
        "row_count": min(len(result_frame), MAX_ROWS_RETURNED),
        "truncated": len(result_frame) > MAX_ROWS_RETURNED,
    }
    if any(b["matched_no_rows"] for b in breakdown):
        empty = [b["label"] for b in breakdown if b["matched_no_rows"]]
        response["warning"] = (
            f"Component(s) {empty} matched zero rows — their contribution is 0. "
            f"Check the transaction-type spelling against the sheet's values."
        )
    if drift:
        response["structure_drift"] = drift
    if region_used:
        response["region_used"] = region_used
    if region_warning:
        response["region_warning"] = region_warning
    return response
