# Data Quality Audit

Checks the workbooks themselves rather than the business. Catches the structural problems that make every other routine quietly wrong.

Most of it is free. The structural half reads only the local graph and makes no API calls at all, so you can run it as often as you like.

**Cadence:** weekly, or after anyone restructures a workbook.
**Cost:** the structural pass is free. The content pass is one call per sheet you choose to sample.
**Output:** a findings list, ordered by how much damage each item can do.

## Prompt, structural pass

This half costs nothing. Run it liberally.

```text
Structural audit of the OneDrive workspace /ERP.
This pass uses only get_workspace_graph. Make no live data calls.

Call get_workspace_graph for /ERP and examine every sheet in every file.
Flag each of the following.

1. Broken header detection. Column names like "Unnamed: 2", "Column1",
   names that are bare numbers, or names that look like dates or data
   values rather than labels. These mean the header row was found in the
   wrong place, so the real headers are being treated as a data row and
   every column label is wrong.

2. Deduplicated columns. Names ending in a numeric suffix that suggest
   two columns shared a header. Worth knowing because conditions must use
   the deduplicated name, not what you see in Excel.

3. Whitespace damage. Leading or trailing spaces in a column name. These
   are invisible in Excel and cause exact-match conditions to fail for
   reasons nobody can see.

4. Near-duplicate names. Two columns in the same sheet whose names differ
   only by case, spacing, or punctuation. A recipe for using the wrong one.

5. Cross-file inconsistency. The same logical column named differently in
   different files: Revenue against Sales against Amount, SKU against
   Item Code. This is what breaks cross_file_aggregate, which matches on
   sheet name and needs the value column to exist under the same name
   everywhere.

6. Sheet naming drift. Sheets that look like they should match across
   files but do not, exactly: trailing spaces, "Q1" against "Q1 2024",
   inconsistent casing.

7. Unroutable sheets. Sheets whose column names are so generic that no
   natural language question could match them. Routing embeds only the
   sheet name, workbook name, and column names, so a sheet with columns
   A, B, C is invisible to the query tool no matter what is in it.

Report findings grouped by severity:

  BREAKS TOTALS      items 1, 5, 6. These produce wrong numbers silently.
  BREAKS FILTERS     items 3, 4. Conditions fail or match the wrong column.
  DEGRADES ROUTING   item 7. query cannot find the sheet.
  WORTH KNOWING      item 2.

For each finding give the file, the sheet, the specific column or name,
and the concrete fix. If a category is empty, say so in one line rather
than omitting it, so I know it was checked.
```

## Prompt, content pass

This half makes live calls. Point it at the sheets that matter rather than the whole workspace.

```text
Content audit of these sheets in /ERP: <list them>.

Fetch each one live and report:

1. Blank rates. Per column, how many rows are blank and what percentage.
   Call out any column above 5 percent that is used for filtering or
   aggregation, because nulls are skipped in a sum and the total comes
   out low with no error.

2. Type inconsistency. Columns holding a mix of numbers and text. The
   usual cause is numbers stored as text, sometimes with thousands
   separators or a currency symbol. Numeric conditions silently match
   nothing against these.

3. Embedded totals rows. Any row that looks like a summary rather than a
   record: a blank identifier with a populated value, a label like TOTAL
   or SUBTOTAL, a value equal to the sum of the rows above it. These
   double every aggregate. Show the row and the condition that would
   exclude it.

4. Date columns held as serials. Numeric columns whose values sit in the
   range you would expect for Excel dates. Say which columns need the
   1899-12-30 conversion.

5. Duplicate keys. If a column looks like an identifier, whether any
   value appears more than once, and how many.

6. Outliers. Values more than three standard deviations from the mean in
   numeric columns, and any negative value in a column that should not
   have one. Report them as candidates, not as errors.

For each finding: file, sheet, column, how many rows affected, and what
it breaks. Do not fix anything. If a response comes back truncated, say
so, because these percentages are then computed over a sample.
```

## Scheduling

Structural pass weekly, since it is free:

```bash
0 6 * * 1 /path/to/venv/bin/python /path/to/structural_audit.py
```

Content pass on demand, or monthly ahead of [month-end reconciliation](month-end-reconciliation.md). Running it the day before close is the difference between finding a totals row on the first of the month and finding it after the numbers went out.

## Why this is worth having

Every failure this catches produces a plausible wrong answer rather than an error. A totals row does not throw, it doubles. A null column does not throw, it undercounts. A trailing space in a header does not throw, it returns an empty result that reads as a legitimate finding of zero.

Errors take care of themselves. Silent wrongness is what needs a routine.

## Tuning it

**Add a diff** by writing findings to a file and comparing against the previous run. New findings are the interesting ones, and a stable list of known-and-accepted issues stops being read after the second week.

**Filter out the known-acceptable.** If a sheet legitimately has 40 percent blanks in an optional notes column, list that exception in the prompt so it stops appearing. An audit that cries wolf gets ignored.

**Run the structural pass in CI** if your workbooks are under any kind of change control. It needs no network and no credentials beyond reading `graph.json`, so it is cheap to gate on.
