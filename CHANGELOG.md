# Changelog

## 0.3.0 — 2026-08-08

### You must re-scan. Tools refuse to run until you do.

`graph.json` gained a per-sheet region map that cannot be back-filled from an
existing graph — the regions come from formulas no earlier version fetched. Any
workspace scanned by 0.2.0 or earlier now fails with a message naming the fix.
Run `scan_workspace` once per folder.

### Why: a sheet is not a table

Measured against a real eight-sheet contract workbook, the 0.2.0 header
heuristic was right on three sheets. On the rest it took the client banner, or
a value, as the header row — and because the graph held exactly one
`header_row` per sheet, every section title, repeated header and totals block
below the first region was handed to the agent as data. `aggregate(sum, 'Qty
(Kgs)')` on one such sheet returned 277,371 against a true figure of 45,000. A
six-fold overcount, silently.

**If you ran a 0.2.0 aggregate over a sheet with more than one table on it, the
number was probably too large. Re-check it.** `get_workspace_graph` now lists
`multi_region_sheets`, which is exactly the set of sheets where this was
possible.

### Table regions, derived from the workbook's own formulas

- A sheet's `SUM`/`SUMIF(S)`/`COUNT`/`COUNTA`/`COUNTIF(S)`/`AVERAGE` calls
  already state where each block starts and ends: `=SUM(C7:C31)` in a totals
  row is the author declaring the extent of the table above it. Scan now reads
  the sheet's formulas, merges the ranges those functions refer to, and records
  the result as `regions` — one entry per table body, in absolute sheet rows.
  On the reference workbook this finds all eleven bodies across the eight
  sheets.
- The header row is taken as `body_start - 1` and outranks the old heuristic,
  which is what moves the workbook from three sheets right to seven. The four
  remaining misses are all exactly one row late, all caused by an
  `Opening Balance` line between the header and the first numbered row. Nothing
  deterministic can resolve those, so every region is marked
  `layout_confidence: "unconfirmed"` and `unclaimed_rows` reports the spans no
  region accounts for, rather than pretending they are data.
- Value sampling is now confined to the first table body, so section titles and
  totals rows below it no longer show up as "values this column contains".

### Cross-sheet references are recorded as facts

`='CMC Statement'!C48` is a stated dependency, not an inference. Those are
extracted at scan and appear in `relationships` with `source: "formula"` and
confidence 1.0. They carry `kind: "reference"` to distinguish them from the
`kind: "join_key"` edges that name-and-value inference produces — a reference
says "this cell reads that cell", not "these columns hold the same entities",
so `join_sheets` explicitly ignores them rather than joining on a meaningless
key.

### Smaller fixes from the same workbook

- Header cells wrapped across lines in Excel (`'Qty\n(Kgs)'`) now collapse to
  `'Qty (Kgs)'`; a column an agent cannot type is a column it cannot query. The
  exact spelling is kept in `columns_raw` when it differs.
- Subscripts and superscripts fold to ASCII digits in name matching, so a query
  saying `H2SO4` reaches `Sulphuric Acid (H₂SO₄)` and the sheet named
  `H2SO4 & Soda Ash`.

### Not in this release

Region-*aware* tool arguments. `filter_sheet` and `aggregate` still read a
sheet as a whole, so on a multi-region sheet you must still narrow with
conditions or use `derive`. Regions are visible in the graph and in
`inspect_file` so an agent can see the trap; making the tools take a `region`
argument, and making `aggregate` refuse on transactional regions, comes next.

## 0.2.0 — 2026-08-08

### If you ran 0.1.0 in production, read these two first

**The server polluted its own protocol stream.** 0.1.0 wrote diagnostics
(`[Embed] No index found...`, retry notices, scan progress) to stdout — the
same stream the MCP stdio transport uses for JSON-RPC frames. Lenient clients
dropped the bad lines; strict clients desynced mid-session. Every runtime
diagnostic now goes to stderr, and a regression test drives a real subprocess
through a full `initialize` → `tools/call` conversation and asserts every
stdout line parses as JSON, so this class of bug cannot ship again.

**`cross_file_aggregate` returned silently partial totals.** Files whose
sheet was named slightly differently (`Sales` vs `Sales 2024` vs `sales `)
were excluded from the total with no warning, no listing, nothing — the
README's "visibly partial" claim was false, and any 0.1.0 total computed over
a workspace with inconsistent sheet naming may have been wrong without any
signal. Excluded files are still excluded (guessing would be worse), but they
are now listed in `unmatched_files` with their actual sheet names and
`did_you_mean` candidates, the `warning` field states the count, and
`get_workspace_graph` reports `sheet_name_variants` so fragmentation is
visible before aggregating. Re-check any recurring 0.1.0 cross-file report
against 0.2.0 output.

### New capability: one-call single-cell lookup

- **`lookup`** answers "what's the contracted rate for Titanium Dioxide under
  the BESTEX contract?" in one tool call: the key/value/return triple is
  resolved from the query against values sampled at scan time (no nested LLM
  call), routing uses value-level evidence, and only the key column plus the
  matched row are read — about 3% of the bytes of a whole-sheet fetch in the
  acceptance fixture. Confidence is explicit: `high` (corroborating sheets
  listed), `ambiguous` (several rows; all returned, value null), `conflict`
  (sheets disagree; every version returned, value null), or `found: false`
  with fuzzy spelling suggestions. No response is ever a bare value.
- **`get_cell`** reads exactly one addressed cell (A1 notation or a
  single-cell named range) in one Graph request.

### New tools

- **`join_sheets`** — live two-sheet join with normalised key matching. When
  keys are omitted it uses a declared or high-confidence inferred
  relationship and refuses (listing candidates) rather than guessing.
- **`derive`** — signed sums over transaction types
  (`receipts − consumption − returns`) computed per grouping key in pandas,
  replacing N filter calls plus mental arithmetic. Components matching zero
  rows are flagged loudly.

### Structure intelligence

- Relationships are now real: inferred at scan time from matching column-name
  tokens plus overlapping sampled values, with a confidence score and the
  evidence. User-declared relationships in `~/.excelmcp/relationships.yaml`
  override inference. (In 0.1.0, `relationships` was advertised but always
  empty.)
- Date columns are detected from `numberFormat` at scan time and served as
  ISO-8601 strings; conditions accept ISO literals
  (`{"Batch Date": ">=2026-01-01"}`). The old RULE 9 told the model to
  convert serials by hand — a known confabulation source, now removed.
- Structure drift is flagged on every live fetch: if the live header row no
  longer matches the scan-time columns, responses carry `structure_drift`
  with both versions. `get_workspace_graph` reports `scan_age`.
- Sheet descriptions now embed sampled values from low-cardinality columns,
  and retrieval is two-stage (vector then lexical rerank), so twenty
  structurally identical workbooks are distinguishable at routing time.
  `query` takes `n_results` and `min_score`, and flags `routing_ambiguous`
  near-ties instead of silently picking.

### Correctness

- Exact string matching strips whitespace and casefolds both sides
  (`"closed"` matches `"Closed "`); `exact_case=true` restores strict
  matching. Zero-row results include `zero_match_diagnostics` with what each
  condition matched alone and up to 20 values actually present.
- The condition grammar gains an object form: `in`, `between`, `is_null`,
  `contains`, and combinable comparison operators. `aggregate` takes a list
  `group_by` and a `having` clause, and no longer drops its `truncated` flag.
- The dead `use_for` field is gone; `inspect_file` reports
  `approx_row_count` labelled `as_of_last_scan` instead of implying a live
  count.
- Fetched frames use the same column-name normalisation as the scan, so blank
  headers are `Column_N` everywhere instead of two different schemes.

### Performance

- Data-path range reads use `$select=values,address,rowCount,columnCount` —
  0.1.0 fetched the full Range resource (text, three formula arrays,
  numberFormat, valueTypes) and read only `values`, roughly a 5–6× bandwidth
  cut on every live call.
- Header detection reads a ten-row window instead of whole sheets; scans read
  sheets concurrently.
- All Graph traffic flows through one shared semaphore
  (`EXCELMCP_MAX_CONCURRENCY`, default 8), held across retry sleeps so a 429
  backs off the whole fleet. Cross-file aggregation folds per-file results
  incrementally instead of concatenating every frame in memory, with mean
  kept row-weighted via (sum, count) pairs.

### What lands on disk (changed)

`graph.json` now also stores, per sheet: the used-range address and
dimensions, per-column date types, and — the one deliberate widening of the
no-cached-data guarantee — `sampled_values`, up to 50 distinct text labels
per low-cardinality column, used for routing, lookup and relationship
inference. Answers are still always served from live fetches. See the
README's "What lands on disk" for the full statement.

## 0.1.0

Initial release.
