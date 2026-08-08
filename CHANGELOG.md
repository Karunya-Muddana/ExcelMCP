# Changelog

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
