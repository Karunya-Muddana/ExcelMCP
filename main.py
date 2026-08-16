import os
from datetime import datetime, timezone
from typing import Any, Optional

from fastmcp import FastMCP

from excelmcp.embeddings import update_embeddings
from excelmcp.lookup import execute_get_cell, execute_lookup
from excelmcp.query_engine import (
    MAX_ROWS_RETURNED,
    execute_aggregate,
    execute_cross_file_aggregate,
    execute_derive,
    execute_fetch_region_rows,
    execute_fetch_sheet_rows,
    execute_filter_sheet,
    execute_inspect_file,
    execute_join_sheets,
    execute_query,
    execute_sheet_layout,
)
from excelmcp.structure import (
    discover_structure,
    load_graph,
    sheet_name_variants,
    workspace_relationships,
)

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
Its default detail='summary' is deliberately compact so
it stays affordable on large workspaces: it carries every
filename, sheet name, column name and column type, plus
the region map for multi-region sheets. Reach for
inspect_file when you need one file in full, and
detail='full' only when you genuinely need sampled
values, sheet descriptions or the raw relationship list.

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

RULE 9 — DATE COLUMNS ARRIVE AS ISO-8601 STRINGS
Columns detected as dates at scan time are converted
from Excel serials to ISO-8601 strings ("2026-03-15")
by the server before you see them. Use ISO literals in
conditions: {"Batch Date": ">=2026-01-01"} works
directly. Never do serial-to-date arithmetic yourself.
If a column looks like raw serials (plain floats near
45000), it was not detected as a date — say so rather
than converting by hand, and suggest scan_workspace.

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

RULE 12 — ONE VALUE MEANS ONE lookup CALL
When the user asks for a single figure — a rate, a
price, one item's stock — call lookup first. It returns
the value with file/sheet/cell provenance in one call.
Check its confidence field, surface conflicts and
ambiguity instead of picking a value, and always cite
where the number came from.

RULE 13 — A SHEET WITH SEVERAL REGIONS IS SEVERAL TABLES
multi_region_sheets in get_workspace_graph lists every
sheet holding more than one table body. On one of these
sheets, filter_sheet, aggregate and derive all REQUIRE
region: omitting it RAISES, naming the available regions,
rather than silently summing or netting blocks that were
never meant to be combined. Each region may carry a label
scanned from its own section banner ("NAPHTHALENE",
"OLEUM 65%") — when present it is named right in that
error, so pick the region by what it is called, not by
guessing from index and row numbers. Call sheet_layout
first when unsure which table answers the question — it
lists every region's body span, label, row count, and
confidence without fetching data. fetch_region_rows
fetches one exact region's raw rows (also refuses to fall
back to the whole sheet); fetch_sheet_rows does the same
but accepts allow_full_sheet_fallback=true for a
deliberate whole-sheet read. filter_sheet, aggregate and
derive accept that same allow_full_sheet_fallback=true
escape hatch — use it only for an explicit, clearly-
labeled whole-sheet result, never as a default.
A region's header_row is now scored against its own data,
not just assumed to sit one row above its body — so it can
land ABOVE the formula-derived body start when a row like
an Opening Balance sits between them. That row (and any
others the header names but the region's own totals never
summed) shows up as the region's pre_body span. When you
pass region to derive, its pre_body rows join the frame
alongside the body — but only a components entry whose
conditions actually match one (e.g. a blank transaction-type
cell) picks it up; without that entry the row sits in the
frame unmatched and contributes nothing, same as before. Add
a components entry with sign 1 or base:true for the opening
balance and its value joins the net; a stock figure derived
from a scoped region without one is a stock figure missing
its opening balance. filter_sheet and aggregate never widen
the frame this way: their region continues to mean exactly
the summed body, nothing more.
region ALSO ACCEPTS A LABEL STRING directly — region=
"naphthalene" resolves the same as region=0 when a region's
scanned label plausibly means that (case/whitespace folded,
then the same fuzzy match sheet-name routing already uses:
containment, small typos). If that matches more than one
region it RAISES rather than picking one; disambiguate with
the index instead. Prefer this over memorising index numbers
when the label is the identifying information you actually have.
CACHED REGIONS CAN GO STALE. A structure_drift block in any
response means the sheet was edited since the last scan.
Region bodies are derived from the sheet's own formulas at
scan time, and an edit that inserts or removes rows inside a
table body never touches the header — so it is invisible to
plain header-drift checking. When structure_drift also
carries region_drift_risk, the sheet's live row count no
longer matches what was scanned, and region boundaries here
may be wrong: say so rather than trusting them, and suggest
scan_workspace before reporting a number from that region.
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


# ---------------------------------------------------------------- graph views
#
# The workspace graph carries two kinds of thing, and only one of them belongs
# in a tool response.
#
#   Structure    — filenames, sheet names, header rows, column names and types.
#                  The model cannot query anything without these.
#   Routing evidence — per-sheet `description` text, `sampled_values`,
#                  `columns_raw`, single-sheet region maps, `layout_confidence`.
#                  These exist so the embeddings can tell twenty structurally
#                  identical workbooks apart. They do their work at routing
#                  time; the model never needs to read them.
#
# Shipping both put a 21-file workspace at ~188k characters — roughly 47k
# tokens for the one call RULE 2 makes mandatory, which overflowed the tool
# result and returned truncated JSON. Structure alone is about a fifth of that.
# Hence: summary by default, everything behind detail='full'.

_SUMMARY_SHEET_KEYS = (
    "header_row",
    "header_source",
    "header_confidence",
    "placeholder_ratio",
    "columns",
    "column_types",
    "used_range_address",
    "row_count",
    "column_count",
    "approx_row_count",
)


def _summarize_sheet(sheet: dict) -> dict:
    """Structure only — but the region map survives where it changes an answer.

    A single-region sheet's region list tells the model nothing it would act on,
    so it collapses to a count. A multi-region sheet is the one case where an
    unqualified aggregate silently sums unrelated blocks, so its regions,
    unclaimed_rows and layout_confidence travel in full even in summary mode.
    """
    out = {k: sheet[k] for k in _SUMMARY_SHEET_KEYS if k in sheet}
    regions = sheet.get("regions") or []
    out["region_count"] = len(regions)
    if len(regions) > 1:
        out["regions"] = regions
        out["unclaimed_rows"] = sheet.get("unclaimed_rows", [])
        out["layout_confidence"] = sheet.get("layout_confidence", "unconfirmed")
    return out


def _summarize_files(files: dict) -> dict:
    return {
        name: {
            "last_scanned": info.get("last_scanned"),
            "sheets": {
                sheet_name: _summarize_sheet(sheet_info)
                for sheet_name, sheet_info in info.get("sheets", {}).items()
            },
        }
        for name, info in files.items()
    }


@mcp.tool(
    description="""
Returns the cached structure of the workspace — every
filename, sheet name, column header, column type and
header row.
INSTANT — makes no API call. Reads from local graph.json.
ALWAYS call this first at session start to orient yourself.
The structure varies for every company — never assume,
always discover.

detail='summary' (the DEFAULT, and what you want almost
always) returns structure only. It is compact enough to
call on a large workspace without burning your context.
detail='full' additionally returns each sheet's sampled
values and description text, plus the full inferred
relationship list with its evidence. These are routing
evidence — the semantic router already uses them on your
behalf — so ask for them only when you specifically need
to inspect what values a column contains.

Also returns:
- sheet_name_variants: groups of sheet names that differ
  only in case or whitespace across files. Check it before
  any cross-file operation, because those match by exact
  sheet name and a variant silently shrinks a total.
- multi_region_sheets: sheets holding more than one table
  body. Each such sheet also keeps its `regions` list (the
  table bodies in absolute sheet rows, derived from the
  sheet's own SUM/COUNT formulas, each with a `label` when
  its own section banner named it — pass that to `region`
  by index or body span, but read the label to know which
  table you are picking), its `unclaimed_rows`, and
  layout_confidence, which reads "unconfirmed" wherever
  the region map came from formulas alone and nothing has
  verified it. A plain aggregate over one of these adds up
  blocks that were never meant to be summed — pick the
  region you mean first.
- relationships_available: how many cross-sheet join keys
  are known. You do not need to read them: call join_sheets
  without left_on/right_on and the server resolves a key
  from this set, refusing with candidates when confidence
  is low. Pass detail='full' to see them.
- scan_age: how old the index is, with a note when it is
  stale enough that renamed sheets may mis-resolve.

Use this before any filter_sheet call when unsure which
file or column to query.
"""
)
async def get_workspace_graph(
    folder_path: Optional[str] = None,
    detail: str = "summary",
) -> dict:
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

    detail = (detail or "summary").strip().lower()
    if detail not in ("summary", "full"):
        return wrap_response(
            {
                "error": f"Unknown detail '{detail}'. Use 'summary' or 'full'.",
            },
            [],
            [],
            0,
        )
    full = detail == "full"

    files = workspace.get("files", {})
    file_names = list(files.keys())
    data = dict(workspace)
    data["detail"] = detail
    if not full:
        data["files"] = _summarize_files(files)

    # Sheet-name fragmentation ('Sales' vs 'sales ' vs 'SALES') silently
    # shrinks cross-file totals; surface it before the agent aggregates.
    # Computed from the unsummarized files — variants live in the keys.
    data["sheet_name_variants"] = sheet_name_variants(files)

    # Declared relationships (relationships.yaml) merged ahead of inferred.
    # join_sheets resolves these server-side, so summary mode reports only that
    # they exist: on a workspace this size the list alone ran to ~75k characters
    # of text the model would never act on directly.
    relationships = workspace_relationships(workspace)
    if full:
        data["relationships"] = relationships
    else:
        data.pop("relationships", None)
        data["relationships_available"] = len(relationships)
        data["relationships_note"] = (
            "Join keys resolve server-side: call join_sheets without "
            "left_on/right_on and it picks from these, refusing with the "
            "candidate list when confidence is low. Pass detail='full' to "
            "read them."
        )

    # Sheets holding more than one table region: an unqualified aggregate over
    # one of these sums blocks that were never meant to be added together.
    data["multi_region_sheets"] = [
        {
            "file": file_name,
            "sheet": sheet_name,
            "regions": len(sheet_info.get("regions", [])),
            "layout_confidence": sheet_info.get("layout_confidence", "unconfirmed"),
        }
        for file_name, file_info in files.items()
        for sheet_name, sheet_info in file_info.get("sheets", {}).items()
        if len(sheet_info.get("regions", [])) > 1
    ]
    scan_age = _scan_age(files)
    if scan_age:
        data["scan_age"] = scan_age
    return wrap_response(data, file_names, [], 0)


def _scan_age(files: dict) -> Optional[dict]:
    """Age of the oldest scan in the workspace — 'scanned 47 days ago' is
    information the agent should act on, not something buried in per-file
    timestamps."""
    stamps = [f.get("last_scanned", "") for f in files.values() if f.get("last_scanned")]
    if not stamps:
        return None
    oldest = min(stamps)
    try:
        scanned_at = datetime.fromisoformat(oldest)
    except ValueError:
        return None
    if scanned_at.tzinfo is None:
        scanned_at = scanned_at.replace(tzinfo=timezone.utc)
    days = max(0, (datetime.now(timezone.utc) - scanned_at).days)
    age: dict[str, Any] = {"oldest_last_scanned": oldest, "days_ago": days}
    if days >= 14:
        age["note"] = (
            f"The structure index is {days} days old. Sheets renamed or "
            "restructured since then will mis-resolve; consider asking the "
            "user whether to run scan_workspace."
        )
    return age


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
Embeds your question, retrieves candidate sheets by
vector similarity, reranks them by lexical overlap with
column names and sampled values, fetches the winners
LIVE from OneDrive, and returns results.
Use for exploratory questions when you do not know which
specific file or sheet contains the answer.
n_results controls how many sheets are fetched (default
5); min_score drops weak matches.
CHECK data.routing: when routing_ambiguous is true the
top candidates scored within a tie margin and the choice
between them is effectively arbitrary — confirm with
inspect_file or ask the user instead of trusting one.
Including a distinctive literal value in the question
(a client name, a material) strongly improves routing.
Response metadata.fetched_at confirms this is live data.
For known file/sheet combinations use filter_sheet instead.
"""
)
async def query(
    question: str,
    folder_path: Optional[str] = None,
    n_results: int = 5,
    min_score: Optional[float] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_query(question, folder, n_results, min_score)
    data: dict[str, Any] = {"results": result.get("results", [])}
    if "routing" in result:
        data["routing"] = result["routing"]
    if "message" in result:
        data["message"] = result["message"]
    return wrap_response(
        data,
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

Condition formats (string form):
  Exact match:      {{"ColumnName": "value"}}
  Contains:         {{"ColumnName": "~value"}}   (literal, not regex)
  Comparisons:      {{"ColumnName": ">100"}}  (also >=, <, <=)
  Date bounds:      {{"Batch Date": ">=2026-01-01"}}  (ISO-8601)

Condition formats (object form, combinable):
  IN list:    {{"Status": {{"in": ["Closed", "Shipped"]}}}}
  Range:      {{"Qty": {{"between": [10, 500]}}}}
  Date range: {{"Batch Date": {{">=": "2026-01-01", "<": "2026-04-01"}}}}
  Null check: {{"Notes": {{"is_null": false}}}}
  Contains:   {{"Name": {{"contains": "oxide"}}}}

Multiple conditions are ANDed together; multiple
operators inside one object are ANDed too.
An unknown column name or operator is an error, not an
empty result.

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

region SCOPES TO ONE TABLE on a multi-region sheet (see
multi_region_sheets / RULE 13). Pass the region's index
(0, 1, ...), its body span as it appears in
get_workspace_graph's regions list ("7:28"), or a label
string ("naphthalene") fuzzy-matched against each region's
scanned section-banner label — ambiguous label matches
raise rather than guess. Applied before conditions. Omit it
on a single-region sheet — no effect there.
On a multi-region sheet, region is EFFECTIVELY REQUIRED:
omitting it RAISES, naming every region — by its label
("NAPHTHALENE") when the sheet's own section banner gave it
one, else its span — rather than silently flattening
several tables into one result. Set
allow_full_sheet_fallback=true only for a deliberate,
explicitly-labeled whole-sheet read — then the call
succeeds and the response carries region_warning instead of
raising. An index out of range or a span matching no region
is always an error listing the real regions, never a silent
fallback. The region actually used comes back as
region_used: {{index, body, row_count, label if any}} — cite it.
filter_sheet's rows never include a region's pre_body rows
(an opening balance the region's own formulas never
summed) — only derive folds those in, and only when scoped
to a region.
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
    region: Optional[Any] = None,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_filter_sheet(
        file_name,
        sheet,
        conditions,
        sort_by,
        limit,
        folder,
        exact_case,
        region,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )
    data = {
        "rows": result["rows"],
        "total_matched": result["total_matched"],
        "truncated": result["truncated"],
    }
    for extra in (
        "zero_match_diagnostics",
        "structure_drift",
        "region_used",
        "region_warning",
    ):
        if extra in result:
            data[extra] = result[extra]
    return wrap_response(
        data,
        [result["file"]],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Returns the sheet's live rows, optionally scoped to one table region.
Use this when the agent needs the raw structure of a complex workbook and
must see the actual first rows before it can reason about subtables,
column groupings or header layout. Default limit=500 returns the first
500 rows of the selected region; set limit=None to fetch the full region.
If the sheet has multiple table regions and no region is passed, the default
behavior is strict: it raises and tells you which region to choose instead of
silently flattening the sheet. Set allow_full_sheet_fallback=True only for a
specific, explicit whole-sheet read that is clearly labeled in the response.
This is a factual raw-data read, not a summary or inferred structure guess.
"""
)
async def fetch_sheet_rows(
    file_name: str,
    sheet: str,
    folder_path: Optional[str] = None,
    limit: Optional[int] = 500,
    region: Optional[Any] = None,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_fetch_sheet_rows(
        file_name,
        sheet,
        folder,
        limit,
        region,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )
    data = {
        "rows": result["rows"],
        "total_matched": result["total_matched"],
        "truncated": result["truncated"],
    }
    for extra in ("structure_drift", "region_used", "region_warning"):
        if extra in result:
            data[extra] = result[extra]
    return wrap_response(
        data,
        [result["file"]],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Returns the sheet layout summary: header row, columns, used range, and each
region's label (when its own section banner named it), body span, header_row,
formula_header_row (the raw formula-derived guess before scoring corrected
it, kept for comparison), pre_body (rows the header names but the region's
own formulas never totalled, such as an opening balance — empty when there
is no gap between header and body), row count, source, and confidence. Use
this to decide whether a sheet is one table or several subtables — and which
one is which by name — before fetching rows or aggregating. This tool is
factual metadata about the workbook, not a guess.
"""
)
async def sheet_layout(
    file_name: str,
    sheet: str,
    folder_path: Optional[str] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_sheet_layout(file_name, sheet, folder)
    return wrap_response(
        {
            "header_row": result["header_row"],
            "columns": result["columns"],
            "column_types": result["column_types"],
            "used_range_address": result["used_range_address"],
            "row_count": result["row_count"],
            "region_count": result["region_count"],
            "regions": result["regions"],
            "unclaimed_rows": result["unclaimed_rows"],
            "layout_confidence": result["layout_confidence"],
            **({"structure_drift": result["structure_drift"]} if "structure_drift" in result else {}),
        },
        [result["file"]],
        [result["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Fetches one exact table region by index or body span, and never falls back to
whole-sheet data. Use this for the strict full-region read when the workbook is
known to have multiple subtables and a specific region is required. Pass region
as either the integer index or the body span such as '7:28'.
"""
)
async def fetch_region_rows(
    file_name: str,
    sheet: str,
    region: Any,
    folder_path: Optional[str] = None,
    limit: Optional[int] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_fetch_region_rows(file_name, sheet, folder, region, limit)
    data = {
        "rows": result["rows"],
        "total_matched": result["total_matched"],
        "truncated": result["truncated"],
        "region_used": result["region_used"],
    }
    for extra in ("structure_drift", "region_warning"):
        if extra in result:
            data[extra] = result[extra]
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
group_by is one column name or a list of them.
conditions filters rows before aggregating (same
grammar as filter_sheet, including the object form).
having filters the AGGREGATED rows afterwards, e.g.
having={"Revenue": ">1000"} keeps only groups whose
aggregate exceeds 1000.
SINGLE FILE ONLY.
For totals across multiple files you MUST use
cross_file_aggregate instead — never use this tool
and then manually add results across files.
Get column names from get_workspace_graph first.
If the sheet appears in multi_region_sheets, the default
behavior is strict: pass region (the region's index, e.g. 0,
or its body span like "7:28" from get_workspace_graph's
regions list — or sheet_layout to see each region's label)
to restrict to the table you actually mean, applied before
conditions and grouping; omitting it raises instead of
silently flattening the sheet, naming each region by its
label when the sheet's own section banner gave it one. The
region actually used comes back as region_used: {index,
body, row_count, label if any}. Like filter_sheet, the
summed rows never include a region's pre_body (an opening
balance the region's own formulas never totalled) — only
derive folds those in. Returns rows plus a truncated flag.
If conditions matched zero rows, zero_match_diagnostics
shows what each condition matched alone and the values
actually present — correct the condition and retry.
"""
)
async def aggregate(
    file_name: str,
    sheet: str,
    group_by: Any,
    value_col: str,
    operation: str,
    folder_path: Optional[str] = None,
    conditions: Optional[dict] = None,
    having: Optional[dict] = None,
    region: Optional[Any] = None,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_aggregate(
        file_name,
        sheet,
        group_by,
        value_col,
        operation,
        conditions,
        folder,
        having,
        region,
        allow_full_sheet_fallback=allow_full_sheet_fallback,
    )
    data = {"rows": result["rows"], "truncated": result["truncated"]}
    for extra in (
        "zero_match_diagnostics",
        "structure_drift",
        "region_used",
        "region_warning",
    ):
        if extra in result:
            data[extra] = result[extra]
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
A total is only meaningful if the column means the same
thing in every file — same unit, same currency. Two files
can share a sheet and column name and still be
incommensurable; check before reporting one number.
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
            "drifted_files": result.get("drifted_files", []),
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
exact column names available in a specific file, and in
preference to get_workspace_graph(detail='full') when
your question concerns one file: the payload is a
fraction of the size.
"""
)
async def inspect_file(file_name: str, folder_path: Optional[str] = None) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_inspect_file(file_name, folder)
    sheets = list(result.get("sheets", {}).keys())
    return wrap_response(result, [file_name], sheets, 0)


@mcp.tool(
    description="""
Joins two sheets LIVE on key columns and returns the
merged rows, with filter_sheet's truncation contract
(total_matched, truncated, limit).
Omit left_on/right_on to let the server pick keys from
the workspace's known relationships — it uses a declared
or high-confidence inferred relationship and REFUSES
with the candidate list when confidence is low, rather
than guessing. The keys actually used and their source
are in data.keys. You do NOT need to fetch the
relationship list first; this resolution is why
get_workspace_graph does not ship it by default.
Key matching is normalised (case, whitespace, 45 vs
45.0); null keys never join. join_type: inner, left,
right, outer. Colliding column names get _left/_right
suffixes.
Use this instead of stitching filter_sheet results
together yourself.
"""
)
async def join_sheets(
    left_file: str,
    left_sheet: str,
    right_file: str,
    right_sheet: str,
    left_on: Optional[str] = None,
    right_on: Optional[str] = None,
    join_type: str = "inner",
    folder_path: Optional[str] = None,
    limit: Optional[int] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_join_sheets(
        folder, left_file, left_sheet, right_file, right_sheet,
        left_on, right_on, join_type, limit,
    )
    data = {
        "rows": result["rows"],
        "keys": result["keys"],
        "join_type": result["join_type"],
        "total_matched": result["total_matched"],
        "truncated": result["truncated"],
    }
    if "drifted_files" in result:
        data["drifted_files"] = result["drifted_files"]
    return wrap_response(
        data,
        [left_file, right_file],
        [result["left"]["sheet"], result["right"]["sheet"]],
        result["row_count"],
    )


@mcp.tool(
    description="""
Computes a NET value over transaction types in one call:
sum of sign * groupwise_sum(quantity_col) across the
given components. This is how a stock figure like
  receipts + purchases − consumption − returns
becomes ONE call with the arithmetic done in pandas,
instead of five filter_sheet calls added up in your
head (which RULE 3 forbids).

components is a list of
  {"conditions": {...same grammar as filter_sheet...},
   "sign": 1 or -1, "label": "receipts"}
conditions (optional) pre-filters the sheet before any
component applies.

THE SIGN IS YOURS TO DECIDE, NOT THE SHEET'S. Outbound
movements — returns, issues, consumption — are very often
stored as POSITIVE quantities and distinguished only by a
type label. Read the transaction-type values before
choosing signs, and never assume the column is already
signed because the numbers look plausible.
The response includes a per-component breakdown with
rows_matched. A component that matched ZERO rows is
flagged and warned about — check the spelling of the
transaction type before trusting the net.

A component can use "base": true instead of "sign" — a
one-off row like an opening balance, added with sign +1.
It still needs "conditions" that match exactly that row
(e.g. a blank transaction-type cell: {"Type": {"is_null":
true}}), and it goes through the same zero-match
flagging as any other component — check rows_matched is 1
before trusting it.

group_by is OPTIONAL. Omit it for one net total across
every matching row (a single row with just "net" — no
group key) — do not pass [] or a guessed column name to
mean "no grouping". Give a column name or list of them
only when you actually want the net broken out per group.

region SCOPES TO ONE TABLE on a multi-region sheet — see
filter_sheet's region parameter for the full contract
(index or body span, or sheet_layout to see each region's
label; applied before conditions and components). On a
multi-region sheet region is EFFECTIVELY REQUIRED: omitting
it RAISES, naming every region by its label when the sheet's
own section banner gave it one, rather than silently netting
several tables together. Set allow_full_sheet_fallback=true
only for a deliberate, explicitly-labeled whole-sheet net —
the response then carries region_warning instead of raising.
The region actually used comes back as region_used: {index,
body, row_count, label if any}. This is what turns "receipts
minus returns across the whole sheet" into "net for
Naphthalene" instead of Naphthalene and Oleum netted
together.

UNLIKE filter_sheet/aggregate, a resolved region's pre_body
rows (an opening balance the region's own formulas never
summed — see sheet_layout) join the frame here too, so a
components entry that matches one (blank transaction-type,
"sign": 1 or "base": true) folds it into the net. This is
what makes derive(region=0, components=[...,
{"conditions": {"Type": {"is_null": true}}, "base": true,
"label": "opening"}]) return the correct stock figure
instead of receipts-minus-returns with the opening balance
quietly missing. region_used carries pre_body: "<span>" when
this happened.
"""
)
async def derive(
    file_name: str,
    sheet: str,
    quantity_col: str,
    components: list,
    group_by: Optional[Any] = None,
    folder_path: Optional[str] = None,
    conditions: Optional[dict] = None,
    region: Optional[Any] = None,
    allow_full_sheet_fallback: bool = False,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_derive(
        folder, file_name, sheet, group_by, quantity_col, components, conditions,
        region, allow_full_sheet_fallback=allow_full_sheet_fallback,
    )
    data = {
        "rows": result["rows"],
        "components": result["components"],
        "truncated": result["truncated"],
    }
    for extra in (
        "warning",
        "structure_drift",
        "region_used",
        "region_warning",
    ):
        if result.get(extra):
            data[extra] = result[extra]
    return wrap_response(
        data, [result["file"]], [result["sheet"]], result["row_count"]
    )


@mcp.tool(
    description="""
Reads EXACTLY ONE cell, LIVE, by address. One Graph
request, tiny payload, no ambiguity.
Use when the location is already known — follow-up
questions, scheduled routines, anything where lookup or
filter_sheet already established the address earlier.
address is A1 notation ("B7") or the name of a
workbook-scoped named range that resolves to one cell.
Serial dates arrive converted to ISO-8601; check
resolved_type. A multi-cell address is an error — use
filter_sheet for ranges.
"""
)
async def get_cell(
    file_name: str,
    sheet: str,
    address: str,
    folder_path: Optional[str] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_get_cell(file_name, sheet, address, folder)
    return wrap_response(result, [file_name], [sheet], 1)


@mcp.tool(
    description="""
ONE-CALL semantic lookup: finds a single cell value
anywhere in the workspace and returns it WITH PROVENANCE
(file, sheet, cell address, the matched row) and a
confidence signal. Reads only the key column and the
matched row — never whole sheets.

Two ways to call it:
1. Natural language: lookup(query="contracted rate for
   Titanium Dioxide under the BESTEX contract"). The
   server resolves the key value against values sampled
   at scan time and picks the return column lexically.
   Works best when the query contains a literal value
   that appears in the data (a client, a material).
2. Explicit: lookup(key_column="Material", key_value=
   "Titanium Dioxide", return_column="Contracted Rate").
   Use this when the query form reports it could not
   parse, or for values too rare to be sampled.
scope={"file": ..., "sheet": ...} narrows the search.

READ confidence BEFORE using the value:
  "high"      — single row matched; corroborating sheets
                (if any) agree. provenance.corroborated_by
                lists them.
  "ambiguous" — the key matched SEVERAL ROWS. value is
                null; every row is in alternatives. Never
                pick one silently.
  "conflict"  — several sheets DISAGREE. value is null;
                every version is in alternatives. Surface
                the conflict to the user.
found=false   — key not found; suggestions holds fuzzy
                near-misses (retry with exact spelling),
                or ambiguity explains why routing failed.
NEVER present a value from this tool without citing
provenance.file, provenance.sheet and provenance.cell.
An "ambiguous" or "conflict" result is a CORRECT ANSWER,
not a failure to work around. Report it as such rather
than retrying with a narrower scope until one value
survives.
"""
)
async def lookup(
    query: Optional[str] = None,
    key_column: Optional[str] = None,
    key_value: Optional[str] = None,
    return_column: Optional[str] = None,
    folder_path: Optional[str] = None,
    scope: Optional[dict] = None,
) -> dict:
    folder = _resolve_folder(folder_path)
    result = await execute_lookup(
        query, key_column, key_value, return_column, folder, scope
    )
    searched = result.get("searched") or []
    files_queried = sorted({s["file"] for s in searched if s.get("file")})
    sheets_queried = sorted({s["sheet"] for s in searched if s.get("sheet")})
    if result.get("provenance"):
        files_queried = files_queried or [result["provenance"]["file"]]
        sheets_queried = sheets_queried or [result["provenance"]["sheet"]]
    return wrap_response(result, files_queried, sheets_queried, 0)


def main() -> None:
    # Explicit stdio transport: this server is spawned as a subprocess by the
    # host agent and speaks JSON-RPC over stdin/stdout. Nothing may be printed
    # to stdout outside that protocol or the agent's parser will desync.
    mcp.run(transport="stdio")


run = main

if __name__ == "__main__":
    main()