# Prompt Library

Prompts you can paste into a session that has ExcelMCP attached. They are written in the second person, addressed to the agent, and grouped by what you are trying to get out of it.

Replace `/ERP` with your folder and the filenames with your own. Everything here is deliberately schema free where it can be, so most of it works unchanged.

---

## Orientation

### Map the workspace

```text
Call get_workspace_graph for /ERP and give me a map of what is in there.
For each file, list its sheets, and for each sheet the column names and
roughly what the sheet appears to be about. This is a free local call, so
do not worry about cost. Do not fetch any data yet.
```

### Explain a file you have never seen

```text
Run inspect_file on Inventory.xlsx in /ERP. Walk me through what each sheet
is, what the columns mean, and which sheet you would use to answer a
question about current stock levels. Flag any column name that looks
ambiguous or that you would want me to clarify before you rely on it.
```

### Find the right sheet for a question

```text
I want to know our total revenue for Q1. Do not fetch anything yet. Look at
the workspace graph for /ERP and tell me which file, which sheet, and which
column you would use, and what your second choice would be. Then wait for
me to confirm before making a live call.
```

This one is worth the extra turn on any question where being wrong is expensive. You get to catch a bad column choice before it becomes a number in a report.

---

## Straight answers

### One filtered lookup

```text
In /ERP, from Inventory.xlsx sheet Stock, show me every row where Status is
Low, sorted by quantity ascending. Check the graph first for the exact
column names. If the result comes back truncated, tell me the total matched
count rather than pretending you saw everything.
```

### A grouped total

```text
In /ERP, aggregate Sales.xlsx sheet Q1: group by Region, sum Revenue.
Show the result as a table sorted high to low, and give me the grand total
underneath. Tell me when the data was fetched.
```

### A total across every file

```text
I need total Q1 revenue across all files in /ERP, not just Sales.xlsx.
Use cross_file_aggregate. Show me the per-file breakdown as well as the
total. If any file was skipped or there is a warning in the response, lead
with that before giving me the number.
```

### An exploratory question

```text
Using /ERP, answer this: which products are we losing money on?
Start with the workspace graph so you know what exists. Use query if you
are unsure which sheet is relevant, then follow up with a precise
filter_sheet or aggregate once you have narrowed it down. Show your
reasoning about which sheets you chose and why.
```

---

## Analysis

### Compare two periods

```text
In /ERP, compare Sales.xlsx Q1 against Q2. For each Region, give me Q1
revenue, Q2 revenue, the absolute change, and the percentage change. Sort by
absolute change descending. Two aggregate calls, one per sheet. Do the
subtraction yourself after both come back, and state both fetch timestamps.
```

### Rank and threshold

```text
From /ERP, Sales.xlsx sheet Q1, give me the top 10 products by revenue and
tell me what share of total revenue those 10 represent. You will need the
grand total as well as the top slice, so aggregate first and do the
arithmetic on what comes back rather than filtering twice.
```

### Cross-reference two files

```text
In /ERP, I want the low stock items from Inventory.xlsx sheet Stock joined
against supplier lead times in Suppliers.xlsx. Fetch both, match them on
whatever key column they share, and show me items where stock is low and
lead time is over 14 days. Tell me the join key you used and how many rows
failed to match on either side.
```

The last sentence matters. An unreported join miss is how a "complete" list quietly loses a third of its rows.

### Spot an anomaly

```text
Pull Sales.xlsx sheet Q1 from /ERP and look for anything that does not
belong: negative values in columns that should be positive, dates outside
the quarter, duplicated identifiers, blank required fields, outliers more
than three standard deviations out. Report what you find with the row
values. Do not fix anything, just tell me.
```

---

## Verification

### Double check a number

```text
You told me total Q1 revenue is <number>. Verify it a different way: get the
per-file breakdown with cross_file_aggregate and separately aggregate each
file on its own, then compare. If the two approaches disagree by any amount,
show me both numbers and your best guess at the cause. Do not reconcile them
silently.
```

### Audit a truncated result

```text
That last result had truncated=true. Tell me total_matched, then either
narrow the conditions so the full set fits under the cap, or switch to
aggregate so I get a correct summary instead of an arbitrary first page.
Do not just show me more rows.
```

### Prove the data is live

```text
Read the metadata on that last response. Tell me fetched_at and is_cached.
Then fetch the same sheet again and tell me whether anything changed between
the two calls. I want to see the freshness for myself.
```

---

## Reporting

### Executive summary

```text
Using /ERP, write me a one page summary of the current inventory position.
Cover: total SKUs, how many are low or out of stock, the value at risk, and
the five items that need attention first. Every number must come from a
tool call, and put the fetch timestamp at the top. No estimates, no
placeholders, and if something is unavailable say so instead of filling it in.
```

### Table for pasting elsewhere

```text
From /ERP, give me Q1 revenue by Region as a plain markdown table, columns
Region and Revenue, sorted descending, with a total row at the bottom. No
commentary above or below the table. Round to whole numbers.
```

### Narrative with citations

```text
Answer this from /ERP: how did the Northeast region perform in Q1?
Write it as prose, but after every factual claim put the source in brackets
like [Sales.xlsx / Q1 / Revenue, fetched 14:32Z]. If you cannot source a
claim from a tool response, do not make the claim.
```

---

## Data quality

### Header sanity check

```text
For every sheet in /ERP, look at the column names in the workspace graph and
flag anything suspicious: names like "Unnamed: 3", columns that got
deduplicated with a numeric suffix, names with leading or trailing spaces,
or two columns whose names differ only in case. These usually mean the
header row was detected wrong. Report per sheet. This is all free graph
reads, no live calls needed.
```

### Completeness pass

```text
Fetch Inventory.xlsx sheet Stock from /ERP and tell me, per column, how many
rows are blank and what percentage that is. Then tell me which columns are
complete enough to aggregate on and which are not.
```

### Find the totals rows

```text
Fetch Sales.xlsx sheet Q1 from /ERP and check whether the sheet contains any
summary or subtotal rows mixed in with the data, the kind that would double
count if I aggregated. Show me any row that looks like a total rather than a
record, and tell me what condition would exclude them.
```

---

## Anti-prompts

Things that look reasonable and are not. Each is followed by what to say instead.

**Do not say:** "Read the Excel files in my ERP folder and summarise them."
The agent may try the local filesystem, find nothing, and tell you the folder is empty.
**Say instead:** "Using the ExcelMCP tools on the OneDrive folder /ERP, ..."

**Do not say:** "Add up the revenue from each file and give me the total."
This is an instruction to do mental arithmetic across tool calls, which is the single most common source of wrong totals.
**Say instead:** "Use cross_file_aggregate to total revenue, and show the per-file breakdown alongside it."

**Do not say:** "Scan the workspace and then answer my question."
`scan_workspace` re-crawls every workbook over the API. It is for when files changed, not for the start of a session.
**Say instead:** "Call get_workspace_graph to orient yourself, then answer."

**Do not say:** "What is in the Quantity column?"
There may be no column called Quantity. Naming a column you have not verified invites the agent to pick the closest match and not mention it.
**Say instead:** "Check the graph for the stock quantity column, tell me its exact name, then use it."

**Do not say:** "Give me all the low stock items."
Silently caps at 1000 rows if there are more.
**Say instead:** "Give me the low stock items, and if the result is truncated tell me the total matched count before showing me anything."
