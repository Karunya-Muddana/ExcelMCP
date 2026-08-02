# Month-End Reconciliation

The routine where being wrong is expensive, so it computes every total twice by different routes and refuses to smooth over a disagreement.

**Cadence:** first working day of the month.
**Cost:** highest of the routines here. Touches every file.
**Output:** a reconciliation report that either agrees with itself or says loudly that it does not.

## Prompt

```text
Month-end reconciliation for the OneDrive workspace /ERP,
covering <month>.

Step 1. Call get_workspace_graph for /ERP. List every file that contains
the sheet in scope, and note the exact column names in each. Files may
disagree about column naming. Record the differences before you fetch
anything, and if a file uses a different name for the value column, say so
in the report rather than silently skipping it.

Step 2. Compute the total two ways.

  Route A: cross_file_aggregate on the sheet, summing the value column,
  with whatever conditions scope it to <month>.

  Route B: one aggregate call per file, then add the per-file results
  yourself.

Step 3. Compare A against B.

  If they match exactly, say so in one line and move on.

  If they do not match, that is the finding. Stop summarising and start
  investigating. Report both numbers, the difference, and the per-file
  breakdown from each route side by side so the divergent file is
  visible. Then give your best hypothesis: a file skipped by route A, a
  column named differently in one file, a totals row counted in one route
  and not the other, a row with a null value column. Do not pick whichever
  number looks more plausible.

Step 4. Check the response from route A for skipped_files and warning.
If either is populated, that goes at the very top of the report, above
the totals. A total computed over a subset is not a total, and presenting
it as one is worse than reporting a failure.

Step 5. Write the report:

  ## Month-end reconciliation, <month>
  Fetched <fetched_at>

  **Status:** RECONCILED / DISCREPANCY / INCOMPLETE

  <if INCOMPLETE: which files were skipped and why, first, before
   any number appears in this document.>

  **Total**
  Route A (cross_file_aggregate): <n>
  Route B (sum of per-file):      <n>
  Difference:                     <n>

  **By file**
  <table: file, route A contribution, route B contribution, delta.>

  <if DISCREPANCY: an investigation section. Which file diverges,
   what you think caused it, and the specific check that would
   confirm it.>

  **Data quality notes**
  <bullets: null values in the aggregated column, suspected totals
   rows, files using non-standard column names, anything with
   truncated=true. Omit the section if genuinely empty.>

Rules for this run:
- Never reconcile a difference by adjusting a number. Report it.
- Status is INCOMPLETE if any file was skipped, regardless of whether
  the two routes agree on the files that were read. Two routes agreeing
  over the same incomplete subset is not reassurance.
- Every figure comes from a tool response. No carried-forward numbers,
  no estimates, no filling a gap with last month.
- If truncated=true appears anywhere, the affected total is suspect.
  Say which one.
```

## Scheduling

First working day of the month. Cron cannot express "first working day", so run it on the first three days and have the routine no-op if it already ran, or just accept a weekend run:

```bash
0 9 1 * * /path/to/venv/bin/python /path/to/month_end.py
```

Given the stakes, this is a good candidate for a human trigger rather than a schedule. Run it on demand, read it properly.

## Why two routes

Because the failure mode this catches is invisible to a single route.

`cross_file_aggregate` finds files by sheet name. A workbook where the sheet is called `Q1 ` with a trailing space, or `Q1 Final`, does not match, and route A returns a clean, plausible, wrong total with no error. Route B, driven from the file list in the graph, includes it. The two disagree, and the disagreement is the alarm.

The opposite case also happens: a file with a totals row inflates its own `aggregate` result, so route B is high and route A is right. Either way you learn something you would not have learned from one number.

This is also why the prompt forbids picking a winner. The instinct on seeing 1,204,338 against 1,204,900 is to assume the smaller is missing something and go with the larger. Sometimes the larger is double counting. The routine's job is to surface it, and yours is to decide.

## Tuning it

**Add a prior-month comparison** if you want a sanity check on the magnitude. A total that moved 40 percent month over month is worth a look even when both routes agree.

**Add a third route** for high-stakes numbers: fetch the raw rows with `filter_sheet` and count them, checking the row count against what the aggregates implied. Only practical under the 1000 row cap.

**Scope by conditions, not by sheet,** if your months live in one sheet with a date column rather than one sheet per period. Same structure, different step 2.

## What goes wrong

**Sheet name mismatches are the main event.** Trailing spaces, `Q1` against `Q1 2024`, casing. This routine exists largely to catch them.

**Null values in the aggregated column.** Pandas skips nulls in a sum, so both routes agree and both are low. The data quality section is what surfaces this, which is why it is not optional.

**The 1000 row cap on `aggregate`.** It applies to the grouped output, not the input, so it only bites when you group by something high cardinality like SKU. If you see `truncated` on a month-end aggregate, group by something coarser.

**Files that legitimately should not be included.** An archive workbook with an old `Q1` sheet will be picked up by route A and inflate everything. Either move it out of the folder or add a condition that excludes it, and note the exclusion in the report so next month's reader knows it was deliberate.
