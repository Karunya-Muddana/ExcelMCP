import os
from datetime import datetime, timezone
from typing import Any, Optional

from fastmcp import FastMCP

from excelmcp.embeddings import update_embeddings
from excelmcp.query_engine import (
    MAX_ROWS_RETURNED,
    execute_aggregate,
    execute_cross_file_aggregate,
    execute_filter_sheet,
    execute_inspect_file,
    execute_query,
)
from excelmcp.structure import discover_structure, load_graph, sheet_name_variants

mcp = FastMCP(
    "ExcelMCP",
    instructions="""
Universal live Excel intelligence for AI agents.
Connects to Microsoft OneDrive via Microsoft Graph API.
Works with ANY company's Excel workspace — no schema
configuration required. Structure is discovered
automatically and stored in the workspace graph.

══ CRITICAL RULES — FOLLOW THESE EXACTLY ══

RULE 1 — DATA IS ALWAYS LIVE
Every tool fetches data from OneDrive at the moment it
is called. The metadata.fetched_at field in every
response proves when the fetch happened.
is_cached is ALWAYS false — there is no data cache.
Never assume you already know a value — always fetch it.

RULE 2 — ORIENT YOURSELF FIRST
At the start of every session call get_workspace_graph
to discover what files and sheets exist in this workspace.
The structure is different for every company and every
folder. Never assume file names, sheet names, or column
names — always discover them from the graph first.

RULE 3 — NEVER CALCULATE CROSS-FILE TOTALS IN YOUR HEAD
If a question involves more than one file, you MUST call
cross_file_aggregate to get the total.
Never add up numbers from individual filter_sheet calls
in your head or using Python. Always verify the total
with cross_file_aggregate. If the two numbers differ,
report both and flag the discrepancy.

RULE 4 — NEVER USE LOCAL FILESYSTEM
All Excel files are on OneDrive, not the local machine.
Never call openpyxl, pandas read_excel, or any terminal
command to find or read Excel files. Use these tools only.

RULE 5 — DO NOT CALL scan_workspace AT SESSION START
The workspace is already indexed from setup. Calling
scan_workspace makes many slow API calls unnecessarily.
Use get_workspace_graph instead — it is instant.
Only call scan_workspace when the user explicitly says
new files have been added or sheets have been renamed.

RULE 6 — PYTHON IS FOR MATH ONLY AFTER FETCHING
Python code execution is only permitted to do arithmetic,
date calculations, or formatting on data already returned
by these tools. Never use Python to fetch or parse files.

RULE 7 — COLUMN NAMES ARE UNKNOWN UNTIL DISCOVERED
Every company uses different column names. Never assume
a column is called "MaterialName" or "Quantity" or
"Status". Always call get_workspace_graph or inspect_file
first to see the actual column names in this workspace,
then use those exact names in filter_sheet and aggregate.

RULE 8 — NEVER MERGE DATA ACROSS FILES INCORRECTLY
Each file is a separate entity with its own data scope.
Values with the same name in different files represent
different real-world things. Never add them together
unless the user explicitly asks for a combined total,
and even then use cross_file_aggregate not mental math.

RULE 9 — DATE VALUES MAY BE EXCEL SERIALS
Dates stored as numbers in Excel are serial values.
Convert using Python after fetching if needed:
  from datetime import date, timedelta
  real_date = date(1899, 12, 30) + timedelta(days=serial)

RULE 10 — TRANSACTION-BASED DATA NEEDS TYPE FILTERING
If a workspace has transaction-based tracking (receipts,
consumption, returns etc.), never sum a quantity column
raw. Always filter by transaction type first, then
aggregate. Inspect the data structure before querying.

RULE 11 — RESULTS MAY BE TRUNCATED
Row-returning tools cap output. If a response has
truncated=true, total_matched tells you how many rows
actually matched. Narrow the query with conditions or
use aggregate instead of assuming you saw everything.
""",
)


def _resolve_folder(folder_path: Optional[str]) -> str:
    """Resolves the workspace folder, falling back to the setup-configured default.

    excelmcp-setup writes EXCELMCP_DEFAULT_FOLDER into the MCP server config, but
    nothing read it, so every call had to pass folder_path explicitly.
    """
    resolved = (folder_path or os.environ.get("EXCELMCP_DEFAULT_FOLDER") or "").strip()
    if not resolved:
        raise ValueError(
            "No folder_path given and EXCELMCP_DEFAULT_FOLDER is not set. "
            "Pass the OneDrive folder to query, e.g. folder_path='/ERP'."
        )
    if not resolved.startswith("/"):
        resolved = "/" + resolved
    return resolved.rstrip("/") or "/"


def wrap_response(
    data: Any,
    files_queried: list[str],
    sheets_queried: list[str],
    row_count: int = 0,
) -> dict:
    now = datetime.now(timezone.utc)
    return {
        "data": data,
        "metadata": {
            "source": "live_onedrive_fetch",
            "fetched_at": now.isoformat(),
            "files_queried": files_queried,
            "sheets_queried": sheets_queried,
            "row_count": row_count,
            "is_cached": False,
            "data_freshness": (
                "LIVE — fetched at query time. "
                "This data was never stored or cached. "
                "It reflects OneDrive state at "
                + now.strftime("%Y-%m-%dT%H:%M:%SZ")
            ),
        },
    }


@mcp.tool(
    description="""
Returns the cached file structure — all filenames, sheet
names, and column headers.
INSTANT — makes no API call. Reads from local graph.json.
ALWAYS call this first at session start to orient yourself.
Shows you exactly which files exist, what sheets they have,
and what columns are in each sheet. The structure varies
for every company — never assume, always discover.
Also returns sheet_name_variants: groups of sheet names
that differ only in case or whitespace across files —
check it before any cross-file operation, because those
match by exact sheet name.
Use this before any filter_sheet call when unsure which
file or column to query.
"""
)
async def get_workspace_graph(folder_path: Optional[str] = None) -> dict:
    folder = _resolve_folder(folder_path)
    graph = load_graph()
    workspace = graph.get("workspaces", {}).get(folder)
    if not workspace:
        known = list(graph.get("workspaces", {}).keys())
        return wrap_response(
            {
                "error": f"Workspace '{folder}' not found. Run scan_workspace first.",
                "known_workspaces": known,
            },
            [],
            [],
            0,
        )
    file_names = list(workspace.get("files", {}).keys())
    data = dict(workspace)
    # Sheet-name fragmentation ('Sales' vs 'sales ' vs 'SALES') silently
    # shrinks cross-file totals; surface it before the agent aggregates.
    data["sheet_name_variants"] = sheet_name_variants(workspace.get("files", {}))
    return wrap_response(data, file_names, [], 0)


@mcp.tool(
    description="""
Rescans the OneDrive folder and rebuilds the structure
index and embeddings. SLOW — makes many API calls.
ONLY call when: new .xlsx files have been added to OneDrive,
or existing sheet names or column headers have changed.
DO NOT call this at session start.
DO NOT call this before every query.
The workspace is already indexed from setup.
Use get_workspace_graph for instant structure access.
"""
)
async def scan_workspace(folder_path: Optional[str] = None) -> dict:
    folder = _resolve_folder(folder_path)
    workspace = await discover_structure(folder)
    file_names = list(workspace["files"].keys())
    if not file_names:
        return wrap_response(
            {
                "message": f"No .xlsx files found in '{folder}'.",
                "files_found": 0,
                "files": [],
            },
            [],
            [],
            0,
        )
    update_embeddings(folder, workspace["files"])
    return wrap_response(
        {
            "message": f"Scanned workspace '{folder}'.",
            "files_found": len(file_names),
            "files": file_names,
        },
        file_names,
        [],
        0,
    )


@mcp.tool(
    description="""
Natural language question with automatic RAG routing.
Embeds your question, finds the most relevant sheets via
semantic similarity search across the workspace graph,
fetches them LIVE from OneDrive, returns results.
Use for exploratory questions when you do not know which
specific file or sheet contains the answer.
Works universally — no knowledge of the schema needed.
Response metadata.fetched_at confirms this is live data.
For known file/sheet combinations use filter_sheet instead.
"""
)
async def query(question: str, folder_path: Optional[str] = None) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_query(question, folder)
    return wrap_response(
        result.get("results", result),
        result.get("files", []),
        result.get("sheets", []),
        result.get("row_count", 0),
    )


@mcp.tool(
    description=f"""
Fetches a specific sheet LIVE from OneDrive and returns
rows matching the given conditions. Always live — no cache.
Use when you already know which file and sheet to query.
Get column names from get_workspace_graph first.

Condition formats:
  Exact match:      {{"ColumnName": "value"}}
  Contains:         {{"ColumnName": "~value"}}   (literal, not regex)
  Greater than:     {{"ColumnName": ">100"}}
  Greater or equal: {{"ColumnName": ">=100"}}
  Less than:        {{"ColumnName": "<100"}}
  Less or equal:    {{"ColumnName": "<=100"}}

Multiple conditions are ANDed together.
An unknown column name is an error, not an empty result.

MATCHING IS NORMALISED, NOT STRICT: exact string matches
ignore case and surrounding whitespace ("closed" matches
"Closed "), because Excel cells carry stray whitespace
constantly. Pass exact_case=true for byte-for-byte
matching. Contains (~) is case-insensitive.
If zero rows match, the response includes
zero_match_diagnostics showing what each condition
matched on its own and the values actually present in
the column — use it to correct a near-miss and retry
instead of concluding the data does not exist.
At most {MAX_ROWS_RETURNED} rows are returned; check the
truncated and total_matched fields in the response.
"""
)
async def filter_sheet(
    file_name: str,
    sheet: str,
    conditions: dict,
    folder_path: Optional[str] = None,
    sort_by: Optional[str] = None,
    limit: Optional[int] = None,
    exact_case: bool = False,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_filter_sheet(
        file_name, sheet, conditions, sort_by, limit, folder, exact_case
    )
    data = {
        "rows": result["rows"],
        "total_matched": result["total_matched"],
        "truncated": result["truncated"],
    }
    if "zero_match_diagnostics" in result:
        data["zero_match_diagnostics"] = result["zero_match_diagnostics"]
    return wrap_response(
        data,
        [result["file"]],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Fetches a sheet LIVE and runs a grouped aggregation.
Operations: sum, count, mean, min, max.
SINGLE FILE ONLY.
For totals across multiple files you MUST use
cross_file_aggregate instead — never use this tool
and then manually add results across files.
Get column names from get_workspace_graph first.
Returns rows plus a truncated flag. If conditions
matched zero rows, zero_match_diagnostics shows what
each condition matched alone and the values actually
present — correct the condition and retry.
"""
)
async def aggregate(
    file_name: str,
    sheet: str,
    group_by: str,
    value_col: str,
    operation: str,
    folder_path: Optional[str] = None,
    conditions: Optional[dict] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_aggregate(
        file_name, sheet, group_by, value_col, operation, conditions, folder
    )
    data = {"rows": result["rows"], "truncated": result["truncated"]}
    if "zero_match_diagnostics" in result:
        data["zero_match_diagnostics"] = result["zero_match_diagnostics"]
    return wrap_response(
        data,
        [result["file"]],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
MANDATORY for any total spanning more than one file.
Fetches relevant sheets from ALL files in PARALLEL,
applies filter conditions, returns the aggregate total.

WHEN YOU MUST CALL THIS:
- Any total, sum, count, or average across multiple files
- Any cross-file comparison or consolidation
- Verifying a total you calculated from individual files

NEVER calculate cross-file totals by:
- Adding individual filter_sheet results in your head
- Using Python to sum numbers from separate tool calls
- Guessing based on partial data

Always call this AND show per-file breakdown so the
user can verify both agree. If they differ, flag it.

ONLY files whose sheet is named EXACTLY `sheet` are
included in the total. Files without that exact sheet
are listed in unmatched_files, with their actual sheet
names and did_you_mean candidates — they are NEVER
silently included. If the response has a warning,
skipped_files, or unmatched_files, surface that to the
user: the total may be incomplete. Check
sheet_name_variants in get_workspace_graph first to see
naming fragmentation before aggregating.
"""
)
async def cross_file_aggregate(
    sheet: str,
    value_col: str,
    operation: str,
    folder_path: Optional[str] = None,
    conditions: Optional[dict] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_cross_file_aggregate(
        folder, sheet, value_col, operation, conditions
    )
    return wrap_response(
        {
            "total": result["total"],
            "operation": result["operation"],
            "column": result["column"],
            "per_file": result["per_file"],
            "skipped_files": result.get("skipped_files", []),
            "unmatched_files": result.get("unmatched_files", []),
            "warning": result.get("warning"),
        },
        result["files"],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Returns structural metadata for one specific file —
sheet names, column headers, and each sheet's
approx_row_count AS OF THE LAST SCAN (this tool makes
no API call, so the count is not live; treat it as an
order-of-magnitude hint, not a current figure).
INSTANT — reads from cached graph.json.
Use before filter_sheet when you need to confirm the
exact column names available in a specific file.
"""
)
async def inspect_file(file_name: str, folder_path: Optional[str] = None) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_inspect_file(file_name, folder)
    sheets = list(result.get("sheets", {}).keys())
    return wrap_response(result, [file_name], sheets, 0)


def main() -> None:
    # Explicit stdio transport: this server is spawned as a subprocess by the
    # host agent and speaks JSON-RPC over stdin/stdout. Nothing may be printed
    # to stdout outside that protocol or the agent's parser will desync.
    mcp.run(transport="stdio")


run = main

if __name__ == "__main__":
    main()
