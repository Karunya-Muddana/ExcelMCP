# Daily Inventory Check

Runs every weekday morning and tells you what needs attention before anyone asks. The point is that the numbers are fetched at the moment the report is written, not overnight, so a stock movement at 07:15 is in the 07:30 report.

**Cadence:** weekdays, early.
**Cost:** two or three live calls.
**Output:** short markdown, suitable for Slack or email.

## Prompt

```text
Daily inventory check for the OneDrive workspace /ERP.

Step 1. Call get_workspace_graph for /ERP. This is free and instant. Use it
to find the sheet holding current stock levels and note the exact column
names for quantity, reorder point, and status. Do not assume these names.
If you cannot identify a stock sheet, stop and report that instead of
guessing at a different sheet.

Step 2. Fetch the stock sheet live and identify three groups:
  - Out of stock: quantity at or below zero.
  - Low: quantity at or below the reorder point but above zero.
  - Newly low: in the low group, and not in yesterday's list if you have it.

Step 3. Write the report in exactly this shape:

  ## Inventory check, <date>
  Fetched <fetched_at from the tool response>

  **Out of stock: <n>**
  <table: item, location, days out if the data supports it. Max 15 rows.>

  **Low stock: <n>**
  <table: item, on hand, reorder point, shortfall. Max 15 rows,
   sorted by shortfall descending.>

  **Nothing else needs attention.**
  <or omit this line if the above is non-empty>

Rules for this run:
- If either table would exceed 15 rows, show the worst 15 and state the
  full count above the table.
- If a response comes back with truncated=true, say so in the report. Do
  not present a truncated list as complete.
- If a column you need does not exist, name the column you looked for, list
  what does exist, and stop. Do not substitute a similar column.
- No estimates and no placeholders. Every number comes from a tool call.
- If nothing is low and nothing is out, say so in one line. Do not pad it.
```

## Scheduling

```
/schedule
```

Weekdays at 07:30. Paste the prompt above.

Or with cron, driving the server directly:

```bash
30 7 * * 1-5 /path/to/venv/bin/python /path/to/inventory_check.py
```

## Tuning it

**If you have no status column,** the reorder point comparison is doing all the work, and that is fine. Delete the mention of status from step 1.

**If stock lives across several workbooks,** by warehouse or by region, add a step: use `cross_file_aggregate` on the stock sheet to get the total position, then report per file. Keep the per-file breakdown in the output, because that is what tells you which site to call.

**If you want a diff against yesterday,** the agent has no memory between scheduled runs. Have the routine write its output to a file and read the previous one at the start. The "newly low" group only works if you do this.

**To trigger on condition rather than on schedule,** add a first line: "If nothing is out of stock and fewer than three items are low, reply with exactly OK and nothing else." Then filter on that downstream so you only get a message when there is something to see.

## What goes wrong

**Days-out is usually not available.** Most stock sheets hold a current position, not a history, so there is nothing to compute a duration from. The prompt says "if the data supports it" for that reason. If you want it, it has to come from a movements log, which is a different and heavier routine.

**Quantity stored as text.** A column with values like `"12"` or `"1,200"` will not compare numerically. You will see this as an empty low-stock list on a day when you know there are low items. Fix it in the workbook.

**Reorder point blank for some rows.** Those rows silently drop out of the low group. Worth adding a line to the prompt asking it to count rows with a missing reorder point and mention the count, so the gap is visible.
