# Getting Started

A first session that proves the whole chain works, from auth through to a live number. Budget about ten minutes, most of it waiting for the initial scan.

## Before you begin

You need the package installed and `excelmcp-setup` run at least once. If you have not done that, go through the [root README](../../README.md) first and come back.

## Step 1: confirm the server is registered

```bash
excelmcp-setup doctor
```

This checks the pieces independently: whether the package is importable, whether the console script is on PATH, whether a token exists and is still valid, whether a workspace has been scanned, and which agent configs currently reference ExcelMCP.

Then confirm your host can actually see it. In Claude Code:

```
/mcp
```

You want ExcelMCP listed with its seven tools. If the server is registered but shows as failed, jump to [troubleshooting](troubleshooting.md).

## Step 2: orient, without spending anything

Start with the free call. This reads a local JSON file and returns instantly.

```text
Call get_workspace_graph for /ERP and show me every file, every sheet, and
the column names in each. Do not fetch any data.
```

Read the output properly before moving on. You are checking three things.

**Are all your files there?** A missing workbook usually means it was not `.xlsx`, or it lives in a subfolder that was not crawled.

**Are the sheet names right?** Sheets with no detectable columns are skipped during scanning, so a genuinely empty sheet will be absent and that is correct.

**Do the column names look like column names?** This is the important one. If you see `Unnamed: 2` or `Column1` or a row of dates where headers should be, the header row was detected wrong for that sheet. Header detection scans the first several rows for the first one that looks like a header, and it loses to sheets that open with a title block or merged cells. Note which sheets are affected. You can still query them, you just need to know that what the tool thinks is a header is really data.

## Step 3: one live call

Now something that touches the network.

```text
From /ERP, fetch the first few rows of <a sheet from step 2> using
filter_sheet with empty conditions, and show me the metadata block.
```

Check the metadata. `source` should be `live_onedrive_fetch`, `is_cached` should be `false`, and `fetched_at` should be a timestamp from a few seconds ago. That block is on every data response and it is how you tell, at a glance, that you are not looking at something remembered from earlier in the conversation.

## Step 4: prove it is actually live

Worth doing once, because it is the entire premise of the project.

1. Ask the agent for a specific cell value, something like the quantity for one named item.
2. Open that workbook in Excel or the browser, change the value, and save. OneDrive sync takes a few seconds.
3. Ask the exact same question again.

The number should change. If it does not, wait a moment for OneDrive to settle and ask a third time. There is no cache in ExcelMCP to clear, so a stale answer means either the sync has not landed or the agent is answering from its own conversation history rather than making a fresh call. If you suspect the latter, tell it explicitly to re-fetch.

## Step 5: a real question

```text
Using /ERP, answer this: <a question you actually care about>.
Start with the workspace graph, tell me which sheets you chose and why,
then make the live calls and show me the fetch timestamps.
```

Asking it to narrate the sheet choice is not just for show. It is how you catch the agent quietly picking the wrong column in the first session rather than in a report three weeks later.

## What good looks like

A well behaved session has a shape:

```
get_workspace_graph          free, instant, once at the start
inspect_file                 free, when narrowing to one file
filter_sheet / aggregate     one live call, precisely targeted
```

A session that is going badly looks like repeated `query` calls, or `filter_sheet` failing on a column name and being retried with a guess, or `scan_workspace` at the start of every conversation. All three are fixable with the [system prompt](../system-prompt.md).

## Cost and latency, roughly

| Call | Network | Feels like |
|---|---|---|
| `get_workspace_graph` | none | instant |
| `inspect_file` | none | instant |
| `filter_sheet`, `aggregate` | one sheet | under a second to a few seconds |
| `query` | three sheets in parallel | a few seconds |
| `cross_file_aggregate` | every matching file, parallel | scales with file count |
| `scan_workspace` | every sheet of every file | tens of seconds to minutes |

The first embedding call in a fresh install also downloads the `BAAI/bge-small-en-v1.5` model, which is a one time cost of a few seconds and about 130MB on disk.

## Limits worth knowing now

`filter_sheet` and `aggregate` return at most 1000 rows and set `truncated: true` plus `total_matched` when there were more. `query` is tighter: it routes to the 3 most relevant sheets and returns 50 rows from each, because it is meant for exploration rather than extraction. If you need every row, filter it down or aggregate it.

## Next

- [prompts.md](../prompts.md) for prompts by task
- [query-patterns.md](query-patterns.md) for how to phrase things well
- [routines/](../routines/) once you want this happening on a schedule
