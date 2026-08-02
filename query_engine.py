import asyncio
import math
from typing import Any, Optional

import pandas as pd

from excelmcp.embeddings import search
from excelmcp.graph_client import GraphAPIError, get_used_range
from excelmcp.structure import load_graph

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


def _apply_conditions(df: pd.DataFrame, conditions: dict[str, Any]) -> pd.DataFrame:
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
                    df = df[df[col].astype(str) == str(val)]
            except (ValueError, TypeError):
                df = df[df[col].astype(str) == str(val)]

    # Filtering yields views; downstream code assigns to columns (to_numeric
    # coercion), which raises SettingWithCopyWarning and may not propagate.
    return df.copy()


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
        "sheets": {
            name: {
                "header_row": s.get("header_row", 1),
                "columns": s.get("columns", []),
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
) -> dict:
    folder_path = folder_path.rstrip("/")
    effective_limit = _clamp_limit(limit)
    _, file_info = _get_file_info(folder_path, file_name)
    header_row = _resolve_sheet(file_info, sheet_name, file_name)
    df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)

    if conditions:
        df = _apply_conditions(df, conditions)

    if sort_by:
        if sort_by not in df.columns:
            raise ValueError(
                f"sort_by column '{sort_by}' not found. Available: {list(df.columns)}"
            )
        df = df.sort_values(by=sort_by)

    total_matched = len(df)
    result_df = df.head(effective_limit)
    return {
        "rows": _records(result_df),
        "file": file_name,
        "sheet": sheet_name,
        "row_count": len(result_df),
        "total_matched": total_matched,
        "truncated": total_matched > len(result_df),
    }


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
    df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)

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
    return {
        "rows": _records(result.head(MAX_ROWS_RETURNED)),
        "file": file_name,
        "sheet": sheet_name,
        "row_count": min(len(result), MAX_ROWS_RETURNED),
        "truncated": len(result) > MAX_ROWS_RETURNED,
    }


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

    matching = [
        (fn, fi)
        for fn, fi in workspace.get("files", {}).items()
        if sheet_name in fi.get("sheets", {})
    ]
    if not matching:
        raise ValueError(
            f"No files found with sheet '{sheet_name}' in workspace '{folder_path}'."
        )

    ops_allowed = {"sum", "mean", "count", "min", "max"}
    if operation not in ops_allowed:
        raise ValueError(
            f"Unknown operation '{operation}'. Use: {', '.join(sorted(ops_allowed))}."
        )

    async def fetch_file(file_name: str, file_info: dict) -> tuple[str, pd.DataFrame]:
        header_row = file_info["sheets"][sheet_name].get("header_row", 1)
        df = await _fetch_sheet_data(file_info["item_id"], sheet_name, header_row)
        if conditions:
            df = _apply_conditions(df, conditions)
        return file_name, df

    filenames = [fn for fn, _ in matching]
    raw_results = await asyncio.gather(
        *[fetch_file(fn, fi) for fn, fi in matching], return_exceptions=True
    )

    successes: list[tuple[str, pd.DataFrame]] = []
    failures: list[dict[str, str]] = []
    for fname, result in zip(filenames, raw_results):
        if isinstance(result, Exception):
            failures.append({"file": fname, "error": str(result)})
            print(f"[Warning] Skipped {fname}: {result}")
        else:
            successes.append(result)

    if not successes:
        raise GraphAPIError(
            "All files failed to fetch. Cannot aggregate. "
            "Check OneDrive connectivity and try again. "
            + "; ".join(f"{f['file']}: {f['error']}" for f in failures)
        )

    per_file: dict[str, Any] = {}
    frames: list[pd.DataFrame] = []
    missing_col: list[str] = []
    for file_name, df in successes:
        if value_col not in df.columns:
            missing_col.append(file_name)
            continue
        if operation != "count":
            df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
        file_val = df[value_col].agg(operation)
        per_file[file_name] = float(file_val) if pd.notna(file_val) else None
        frames.append(df)

    if not frames:
        raise ValueError(
            f"Column '{value_col}' not found in any file with sheet '{sheet_name}'."
        )

    combined = pd.concat(frames, ignore_index=True)
    if operation != "count":
        combined[value_col] = pd.to_numeric(combined[value_col], errors="coerce")

    if operation == "mean":
        # Averaging per-file means would weight a 3-row file the same as a
        # 30,000-row one. Aggregating the concatenated frame keeps it row-weighted.
        total = combined[value_col].mean()
    else:
        total = combined[value_col].agg(operation)

    warnings: list[str] = []
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
        "total": float(total) if pd.notna(total) else None,
        "operation": operation,
        "column": value_col,
        "per_file": per_file,
        "files": [fn for fn, _ in successes if fn not in missing_col],
        "sheet": sheet_name,
        "row_count": len(combined),
        "skipped_files": failures,
        "warning": " ".join(warnings) if warnings else None,
    }
