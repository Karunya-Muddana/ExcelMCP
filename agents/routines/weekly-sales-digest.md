# Weekly Sales Digest

Monday morning summary of the week just finished, compared against the week before. Mostly aggregations, so it stays cheap even over large sheets.

**Cadence:** Monday morning.
**Cost:** two to four live calls, all aggregates.
**Output:** a page of markdown with tables.

## Prompt

```text
Weekly sales digest for the OneDrive workspace /ERP.

Step 1. Call get_workspace_graph for /ERP. Identify the sales data and note
the exact names of the date column, the revenue column, and any region or
product columns. Do not assume any of these names. If there is no date
column you can filter on, say so and stop, because week over week is not
computable without one.

Step 2. Establish the two windows. Last week is the seven days ending
yesterday. The prior week is the seven days before that. State both date
ranges explicitly at the top of the report so I can check them. If dates
come back as bare numbers they are Excel serials from the 1899-12-30 epoch,
so convert before comparing.

Step 3. For each window, aggregate revenue by region, and separately by
product. Use the aggregate tool. Do not fetch raw rows and total them
yourself.

Step 4. Write the report:

  ## Sales digest, week ending <date>
  Fetched <fetched_at>
  Comparing <last week range> against <prior week range>

  **Headline**
  <two sentences. Total revenue, the change, and the single most
   important driver of that change.>

  **By region**
  <table: region, last week, prior week, change, percent change.
   Sorted by absolute change descending.>

  **Top movers**
  <table: the 5 products up the most and the 5 down the most,
   by absolute revenue change.>

  **Worth a look**
  <bullets. Anything anomalous: a region at zero that was not,
   a product that appeared or vanished, a number that looks
   like a data problem rather than a business event. Omit this
   section entirely if there is nothing.>

Rules for this run:
- Percentage change against a prior value of zero is undefined. Write
  "new" rather than a percentage or an infinity.
- Do the subtraction and the percentages yourself on the numbers the
  aggregate calls returned. That is arithmetic on fetched data and is
  fine. Do not estimate any input to it.
- If a response has truncated=true, say so before the affected table.
- Round money to whole units. Round percentages to one decimal.
- If either window has no rows at all, lead with that. An empty week is
  much more likely to be a data problem than a real result.
```

## Scheduling

Monday at 08:00. Via `/schedule`, or:

```bash
0 8 * * 1 /path/to/venv/bin/python /path/to/weekly_digest.py
```

## Tuning it

**If your sales data is one sheet per quarter,** the date filter still works within a quarter, but a week spanning a quarter boundary lives in two sheets. Add a step telling it to check whether the window crosses a boundary and aggregate both sheets if so. This will bite you four times a year, always at the worst moment.

**If sales are split across multiple workbooks,** swap the per-file `aggregate` calls for `cross_file_aggregate` on the sheet name, and keep the per-file breakdown in the output.

**For a monthly version,** change the windows to calendar months and add a year-over-year column alongside month-over-month. Seasonal businesses find the year-over-year comparison much more informative than the sequential one.

**To make it shorter,** drop the top movers table and cap the headline at one sentence. The digest people actually read is usually four lines.

## What goes wrong

**Date filtering is the fragile part.** There is no date-range condition in the tool set: conditions are ANDed and you cannot put two on the same column. In practice the agent has to fetch a superset and window it after the fact, which means the 1000 row cap can bite on a busy sheet. If the sheet is large, filter on something else selective first, or aggregate by a period column if the sheet has one.

**Percent change on small bases.** A region going from 40 to 120 is a 200 percent increase and probably noise. Consider adding a floor to the prompt: ignore percentage changes where the prior value is under some threshold.

**Timezone drift.** "Yesterday" is evaluated wherever the job runs. For a cloud scheduled agent that may not be your timezone, and a job at 08:00 local can straddle a day boundary. If the boundaries matter, state the dates literally in the prompt rather than saying "last week".
