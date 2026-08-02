# Query Patterns

How to get the right answer out, and why the wrong one comes out when it does.

## Choosing a tool

The decision is almost always about how much you already know.

| You know | Use | Cost |
|---|---|---|
| Nothing about the workspace | `get_workspace_graph` | free |
| The file, want its shape | `inspect_file` | free |
| Roughly what you want, not where | `query` | 3 sheets |
| File, sheet, and columns | `filter_sheet` | 1 sheet |
| File and sheet, want a summary | `aggregate` | 1 sheet |
| The sheet name, want it across all files | `cross_file_aggregate` | all matching files |

`query` is the convenient one and the one to reach for least. It embeds your question, matches it against sheet descriptions by cosine similarity, takes the top 3 sheets, and fetches 50 rows from each. That is a good way to find out where something lives and a poor way to get a number, because 50 rows is a sample and three sheets may include one that is irrelevant.

The strong pattern is two steps: `query` once to locate, then `filter_sheet` or `aggregate` to actually answer. After the first time, you know where the data is, so skip straight to step two.

## How routing actually works

Sheet descriptions are generated at scan time and look like this:

```
Data sheet 'Stock' in workbook 'Inventory.xlsx'. Contains columns: SKU,
Description, Qty On Hand, Reorder Point, Status, Location.
```

That string, and only that string, is what gets embedded. No cell values are involved, which is the whole point, but it has a consequence worth internalising: **routing quality depends entirely on your column names.**

A sheet with columns `SKU, Description, Qty On Hand` will match a question about stock levels. A sheet with columns `A, B, C, Col4` will match nothing, ever, no matter how good the data is. If `query` keeps picking the wrong sheet, the fix is usually to rename the headers in the workbook and rescan, not to reword the question.

Phrasing that helps routing:

- Use the vocabulary your columns use. Ask about "revenue" if the column is Revenue, not "sales value".
- Name the domain. "inventory stock levels" beats "how much do we have".
- Skip the pleasantries. The whole question gets embedded, so "could you possibly tell me" is noise in the vector.

## Conditions

`filter_sheet` takes a dict of column name to condition string. All conditions are ANDed.

| Form | Meaning | Note |
|---|---|---|
| `{"Col": "value"}` | exact match | case sensitive |
| `{"Col": "~value"}` | contains | literal substring, not a regex |
| `{"Col": ">100"}` | greater than | numeric |
| `{"Col": ">=100"}` | greater or equal | numeric |
| `{"Col": "<100"}` | less than | numeric |
| `{"Col": "<=100"}` | less or equal | numeric |

Several things that are not supported, and their workarounds:

**No OR.** Conditions are always ANDed. For "Low or Critical", make two calls and combine, or aggregate on the status column and read both counts off the result.

**No ranges in one condition.** For "between 10 and 50" you cannot write it as one term, because two conditions on the same column would need the same dict key. Filter on one bound with the tool and narrow the other side after the rows come back.

**No negation.** There is no "not equal". Filter positively where you can, or fetch and exclude afterwards.

**A missing column is an error, not an empty result.** This is deliberate and it is one of the better properties of the tool. An empty result set is indistinguishable from "your filter matched nothing", so a typo'd column name would otherwise look like a legitimate finding of zero rows. Instead you get an error listing the columns that do exist. If you see that error, do not guess at the correct name, read the list.

## Truncation

`filter_sheet` and `aggregate` cap at 1000 rows. `query` returns 50 rows per sheet. Both set `truncated: true` and report `total_matched` when there was more.

Truncation is the quietest way to get a wrong answer, because a truncated response looks exactly like a complete one unless you read the metadata. When you see it, do not ask for more rows. Do one of these instead:

- Narrow the conditions until the real result fits.
- Switch to `aggregate`, which summarises the full sheet even when the grouped output itself would be long.
- Accept it explicitly, and say in the output that it is a sample.

## Cross-file totals

`cross_file_aggregate` takes a sheet name, not a list of sources. It looks for that sheet name in every file in the workspace, fetches all matches in parallel, applies your conditions, and returns a total plus a per-file breakdown.

```python
cross_file_aggregate(
    sheet="Q1",
    value_col="Revenue",
    operation="sum",
    folder_path="/ERP",
    conditions={"Status": "Closed"},
)
```

Two fields in the response matter more than the total. `skipped_files` lists files that could not be read, and `warning` explains why. A total computed over four of five files is not the total. Always check both before quoting the number, and if either is populated, lead with that rather than burying it.

The reason to use this tool rather than several `aggregate` calls plus arithmetic is not convenience, it is that language models are unreliable at multi-step arithmetic and completely reliable at reporting a number a tool handed them. Where accuracy matters, do both and compare: if the tool total and your own sum of the breakdown disagree, something is wrong and you want to know.

## Data shapes that mislead

### Transaction logs

A sheet where each row is a movement, with a type column holding values like RECEIPT, ISSUE, RETURN, ADJUSTMENT. Summing the quantity column across all of it produces a number with no meaning, because receipts and issues point in opposite directions.

Always filter by type first:

```python
aggregate(
    file_name="Movements.xlsx", sheet="Log",
    group_by="SKU", value_col="Quantity", operation="sum",
    conditions={"Type": "RECEIPT"},
    folder_path="/ERP",
)
```

Then do the same for issues, then subtract. Three steps, one correct answer.

### Embedded totals rows

Plenty of human-maintained sheets have a bold TOTAL row at the bottom. It is just another row to the tool, so it gets included in any sum, and the answer comes out exactly double. Check for it once per sheet and exclude it with a condition on whatever column distinguishes it.

### Excel date serials

Dates often come back as bare numbers like `45292`. Excel counts days from an epoch of 1899-12-30:

```python
from datetime import date, timedelta
real_date = date(1899, 12, 30) + timedelta(days=45292)
```

The offset looks wrong by one and is not. Excel treats 1900 as a leap year, which it was not, and 1899-12-30 is the constant that absorbs the error for every date after March 1900.

### Wide sheets

A sheet with one column per month is convenient to read and awkward to aggregate, because "total for the year" means summing across twelve columns rather than grouping one. There is no reshape operation in the tool set. Fetch it and do the arithmetic on the returned rows, or restructure the workbook.

### Merged cells and title blocks

Header detection scans the first several rows for the first row that looks like a header. A sheet that opens with a merged title, a blank row, and then the real headers usually works. A sheet with a multi-row header, where the top row groups the bottom row, does not: you get one of the two rows as headers and the other as a data row. Check the graph output for column names like `Unnamed: 3` or numeric suffixes, which are the signature of this.

### Duplicate column names

Two columns with the same header get deduplicated with a numeric suffix at scan time. The graph shows you the deduplicated names, so use those, not the name as it appears in Excel.

## A worked example

The question: "which suppliers are we at risk with?"

Nothing in that sentence names a file, a sheet, or a column, and "at risk" is not a value in any system. Here is the shape of getting it right.

**Orient, free.** `get_workspace_graph` shows `Inventory.xlsx` with a Stock sheet, `Suppliers.xlsx` with lead times, and `Movements.xlsx` with a transaction log.

**Define the term.** At risk probably means low stock combined with a long lead time. Say so out loud before computing, because if the definition is wrong everything after it is wasted.

**Two live calls.** `filter_sheet` on Stock with `{"Status": "Low"}`, and `filter_sheet` on Suppliers with `{"Lead Time Days": ">14"}`.

**Join on the returned rows.** Match on the supplier ID both sheets carry. Report how many rows failed to match, because an unreported join miss silently shrinks the answer.

**State the definition with the result.** "Seven suppliers are at risk, defined as having at least one low-stock item and a lead time over 14 days, as of 14:32Z. Three inventory rows had a supplier ID not present in Suppliers.xlsx and were excluded."

The last sentence is the one that separates a usable answer from a confident one.
