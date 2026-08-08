import asyncio
import math
from typing import Any, Optional

import pandas as pd

from excelmcp.embeddings import search
from excelmcp.graph_client import GraphAPIError, get_used_range
from excelmcp.storage import log
from excelmcp.structure import fuzzy_name_candidates, load_graph

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


async def _fetch_sheet_data(item_id: str, sheet_name: str, header_row: int = 1) -> pd.DataFrame:
    """Fetches live data from Graph API and returns a DataFrame. Never reads from cache."""
    data = await get_used_range(item_id, sheet_name)
    values = data.get("values", [])
    if not values or len(values) < header_row:
        return pd.DataFrame()

    header_idx = header_row - 1
    headers = _deduplicate_headers(
        [
            str(c).strip() if c is not None and str(c).strip() else f"Col{i}"
            for i, c in enumerate(values[header_idx])
        ]
    )

    data_rows = values[header_idx + 1:]
    padded: list[list[Any]] = []
    for row in data_rows:
        if len(row) < len(headers):
            row = row + [None] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[: len(headers)]
        padded.append(row)

    return pd.DataFrame(padded, columns=headers)


def _deduplicate_headers(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            result.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            result.append(h)
    return result


def _numeric_compare(df: pd.DataFrame, col: str, raw: str, op: str) -> pd.DataFrame:
    """Applies a numeric comparison, raising rather than silently skipping.

    A malformed bound such as `">abc"` previously fell into `except: pass`, so the
    condition was dropped and the caller received the *unfiltered* sheet while
    believing it was filtered — a silently wrong answer, which is worse than an error.
    """
    try:
        bound = float(raw)
    except (TypeError, ValueError):
        raise ValueError(
            f"Condition on column '{col}' uses '{op}' but '{raw}' is not a number. "
            f"Use a numeric bound, e.g. {{'{col}': '{op}100'}}."
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
            try:
                numeric_val = float(val)
                numeric_series = pd.to_numeric(df[col], errors="coerce")
                if not numeric_series.isna().all():
                    df = df[numeric_series == numeric_val]
                else:
                    df = df[_exact_string_match(df[col], val, exact_case)]
            except (ValueError, TypeError):
                df = df[_exact_string_match(df[col], val, exact_case)]

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
    file_info = workspace.get("files", {}).get(file_name)
    if not file_info:
        available = list(workspace.get("files", {}).keys())
        raise ValueError(
            f"File '{file_name}' not found in workspace '{folder_path}'. "
            f"Available: {available}"
        )
    return workspace, file_info


def _resolve_sheet(file_info: dict[str, Any], sheet_name: str, file_name: str) -> int:
    """Returns the header row for a sheet, failing loudly if the sheet is unknown.

    An unknown sheet used to fall through to a default header_row of 1 and then
    produce a confusing Graph 404 instead of naming the valid sheets.
    """
    sheets = file_info.get("sheets", {})
    if sheet_name not in sheets:
        raise ValueError(
            f"Sheet '{sheet_name}' not found in '{file_name}'. "
            f"Available sheets: {list(sheets.keys())}"
        )
    return sheets[sheet_name].get("header_row", 1)


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


async def execute_query(question: str, folder_path: str) -> dict:
    folder_path = folder_path.rstrip("/")
    results = search(folder_path, question, n_results=3)
    if not results:
        return {
            "results": [],
            "files": [],
            "sheets": [],
            "row_count": 0,
            "message": (
                "No relevant sheets found for this workspace. "
                "Run scan_workspace first, or check the folder_path."
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
        header_row = file_info.get("sheets", {}).get(sheet_name, {}).get("header_row", 1)
        df = await _fetch_sheet_data(item_id, sheet_name, header_row)
        return {
            "file": file_name,
            "sheet": sheet_name,
            "score": meta.get("score", 0),
            "rows": _records(df.head(QUERY_ROWS_PER_SHEET)),
            "row_count": len(df),
            "truncated": len(df) > QUERY_ROWS_PER_SHEET,
        }

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
                "columns": s.get("columns", []),
                "approx_row_count": s.get("approx_row_count"),
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
) -> dict:
    folder_path = folder_path.rstrip("/")
    effective_limit = _clamp_limit(limit)
    _, file_info = _get_file_info(folder_path, file_name)
    header_row = _resolve_sheet(file_info, sheet_name, file_name)
    full_df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)

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
    return result


async def execute_aggregate(
    file_name: str,
    sheet_name: str,
    group_by: str,
    value_col: str,
    operation: str,
    conditions: Optional[dict[str, Any]],
    folder_path: str,
) -> dict:
    folder_path = folder_path.rstrip("/")
    _, file_info = _get_file_info(folder_path, file_name)
    header_row = _resolve_sheet(file_info, sheet_name, file_name)
    full_df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)

    df = full_df
    if conditions:
        df = _apply_conditions(df, conditions)

    ops_allowed = {"sum", "mean", "count", "min", "max"}
    if operation not in ops_allowed:
        raise ValueError(
            f"Unknown operation '{operation}'. Use: {', '.join(sorted(ops_allowed))}."
        )
    if value_col not in df.columns:
        raise ValueError(f"Column '{value_col}' not found. Available: {list(df.columns)}")
    if group_by not in df.columns:
        raise ValueError(f"Column '{group_by}' not found. Available: {list(df.columns)}")

    if operation != "count":
        df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

    result = getattr(df.groupby(group_by)[value_col], operation)().reset_index()
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
            header_row = file_info["sheets"][sheet_name].get("header_row", 1)
            df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)
            if conditions:
                df = _apply_conditions(df, conditions)
            return file_name, df, None
        except Exception as exc:
            return file_name, None, exc

    # Results are folded incrementally as each file completes: one frame is
    # held at a time instead of concatenating all of them. For mean, (sum,
    # count) pairs accumulate so the result stays row-weighted — averaging
    # per-file means would weight a 3-row file the same as a 30,000-row one.
    # Concurrency is bounded by the Graph client's shared semaphore.
    per_file: dict[str, Any] = {}
    failures: list[dict[str, str]] = []
    missing_col: list[str] = []
    total_sum = 0.0
    total_count = 0
    running_min: Optional[float] = None
    running_max: Optional[float] = None
    row_count = 0

    for future in asyncio.as_completed(
        [fetch_file(fn, fi) for fn, fi in matching]
    ):
        file_name, df, exc = await future
        if exc is not None:
            failures.append({"file": file_name, "error": str(exc)})
            log(f"[Warning] Skipped {file_name}: {exc}")
            continue
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
        "warning": " ".join(warnings) if warnings else None,
    }
