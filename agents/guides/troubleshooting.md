# Troubleshooting

Symptoms, what they mean, and what to do. Start with the built-in diagnostic:

```bash
excelmcp-setup doctor
```

It checks each layer independently, so its output usually tells you which section below to read.

---

## The server does not appear at all

**No ExcelMCP in `/mcp` or the host's MCP panel.**

Confirm it is registered:

```bash
excelmcp-setup list-agents
```

If your host is listed as detected but not configured, register it:

```bash
excelmcp-setup install --only claude-code
```

If the host is not detected at all, it either is not installed where the wizard looks or is not one of the twelve it knows. Get the config block and paste it manually:

```bash
excelmcp-setup install --dry-run
```

Then restart the host properly. Restarting means quitting completely, not closing the window. On macOS in particular, closing a Claude Desktop window leaves the process running with the old config.

---

## The server is registered but fails to start

This is nearly always PATH. Agents launched from a dock, Start menu, or IDE do not inherit your shell environment, so `command: "excelmcp"` resolves in your terminal and not inside the host.

Test the console script the way the host would see it, from a plain environment rather than your configured shell.

The fix is to bypass the script entirely. Edit the config to use an absolute interpreter path with the module form:

```json
{
  "command": "/absolute/path/to/python",
  "args": ["-m", "excelmcp"]
}
```

Find the right interpreter with `python -c "import sys; print(sys.executable)"` in the environment where you installed the package. `python -m excelmcp` works whenever the package is importable, which makes it strictly more reliable than the console script.

The second possibility is that the package is installed into a different interpreter than the one being invoked. Common with `uv`, with pyenv, and with anything involving more than one virtualenv. Check with `python -c "import excelmcp; print(excelmcp.__file__)"`.

---

## Authentication

### "No token found" or the wizard asks you to sign in again

The token cache lives at `~/.excelmcp/token.json`. Refresh tokens expire, and tenant policy can revoke them early. Re-run:

```bash
excelmcp-setup
```

### Access denied (403)

The message names the missing permission. ExcelMCP requests `Files.Read.All`, and in a managed tenant an administrator may need to consent to it before device-code sign in will grant it. This is a tenant policy question, not a bug: take the app registration and the scope to whoever administers your Microsoft 365.

If you are using your own app registration via `EXCELMCP_CLIENT_ID`, confirm it is configured as a public client with device code flow enabled. A confidential client will not work here, because there is no secret to present.

### 401 after working fine for a while

Handled automatically. Tokens are refreshed proactively when they are within ten minutes of expiry, and a 401 triggers one refresh and retry. If you are seeing 401 surface as an error, the refresh token itself is dead, which means signing in again.

### Signed in as the wrong account

Delete `~/.excelmcp/token.json` and re-run the wizard. Deleting the token affects nothing else, since it holds auth material only.

---

## The workspace looks wrong

### "Workspace '/ERP' not found. Run scan_workspace first."

Either nothing has been scanned, or the folder string does not match what was scanned. The graph is keyed by the exact folder path with trailing slashes stripped, so `/ERP` and `/erp` are different keys. The error response includes `known_workspaces`, which tells you what was actually indexed.

### Files are missing from the graph

Only `.xlsx` is scanned. Legacy `.xls`, `.xlsm` with macros, and CSV are all skipped. Subfolders are not crawled, so a workbook in `/ERP/2024/` will not appear in a scan of `/ERP`.

### Sheets are missing from a file

Sheets with no detectable columns are skipped at scan time. For a genuinely empty sheet that is correct behaviour. For a sheet that clearly has data, it usually means header detection found nothing usable, which points at the next item.

### Column names look like garbage

Names like `Unnamed: 3` or `Column1`, or a row of dates where headers should be, mean the header row was detected wrong.

Detection scans the first several rows and takes the first one that looks like a header. It handles a title row and blank rows above the real headers. It loses to multi-row headers, where one row groups another, and it can lose to merged cells.

Options, roughly in order of how much you will like them: restructure the sheet so the header row is unambiguous, or accept it and treat the misidentified row as data, remembering that the first data row is now missing from results.

Duplicate headers are deduplicated with a numeric suffix. That is normal, and the deduplicated names are what you must use in conditions.

### The graph is stale after editing the workbook

Renaming a sheet or adding a column changes structure, and structure is cached. Rescan:

```text
Call scan_workspace for /ERP.
```

Editing cell values needs no rescan. Values are never cached, so they are current on the next call regardless.

---

## Query results are wrong

### `query` keeps picking the wrong sheet

Routing runs on generated sheet descriptions of the form `Data sheet 'X' in workbook 'Y'. Contains columns: ...`. Only names are embedded, never values, so routing quality is a direct function of how descriptive your column headers are. A sheet whose columns are `A, B, C` cannot be matched by any question.

Two fixes, and the first is better: rename the headers in the workbook and rescan, or skip `query` and target the sheet directly with `filter_sheet`.

### "Column 'X' not found. Available: [...]"

Working as intended. Column names are case sensitive and must match the graph exactly, and the error lists what actually exists. Read the list rather than guessing again. This error exists specifically so that a typo cannot masquerade as a legitimate empty result.

### A filter returns nothing when you expected rows

Check, in order: exact case of the value, leading or trailing whitespace in the cell, whether a numeric comparison is being run against a column stored as text, and whether the conditions are over-constrained. Remember that all conditions are ANDed, so two conditions that are individually reasonable can be jointly impossible.

Fetch the sheet with empty conditions first and look at what the values really are. That is one API call and it settles the question.

### The total is double what it should be

Almost always an embedded totals row being counted as data. Fetch the sheet and look at the last few rows.

### The total is meaningless

If the sheet is a transaction log with a type column, summing the quantity across all types adds receipts to issues. Filter by type first. See [query-patterns.md](query-patterns.md) for the shape of that.

### Numbers in a date column

Excel serials. Convert from the 1899-12-30 epoch:

```python
from datetime import date, timedelta
real_date = date(1899, 12, 30) + timedelta(days=serial)
```

### You are sure you saw fewer rows than exist

Check `truncated` and `total_matched` in the response. `filter_sheet` and `aggregate` cap at 1000 rows, `query` at 50 rows from each of 3 sheets. Narrow the filter or aggregate instead of paging.

---

## Performance

### Everything is slow

`query` fetches three sheets, `cross_file_aggregate` fetches every matching file. Both parallelise, so the cost is roughly the slowest sheet rather than the sum, but a workspace with many large workbooks is genuinely slow to total.

Prefer `filter_sheet` and `aggregate` once you know the target. Use the free structure tools to get to that point.

### Rate limited

Graph returns 429 under load. The client honours `Retry-After` and backs off automatically, and only surfaces the error after retries are exhausted. If you are hitting it repeatedly, something is calling live tools in a loop. The usual culprit is an agent that has not been told the structure tools are free, so it calls `query` to answer questions the graph could have answered for nothing. The [system prompt](../system-prompt.md) fixes this.

### The first call after install takes ages

The embedding model, `BAAI/bge-small-en-v1.5`, downloads on first use. Roughly 130MB, once.

### `scan_workspace` takes minutes

Expected. It touches every sheet of every workbook. It is not meant to run at session start, only when files are added or structure changes.

---

## Getting more detail

Run the server directly in a terminal to see what the host is hiding from you:

```bash
python -m excelmcp
```

It will sit waiting for JSON-RPC on stdin, which is correct. Startup errors, missing dependencies, and import failures show up immediately here and are usually invisible inside a host.

Confirm the offline test suite passes, which isolates logic problems from auth and network problems:

```bash
pytest tests/test_unit.py
```

If tools work here but not in your host, the problem is in the config, and [hosts.md](hosts.md) is the place to look.

## Filing a bug

Issues go to <https://github.com/Karunya-Muddana/ExcelMCP/issues>. Useful things to include: the output of `excelmcp-setup doctor`, your host and OS, the tool call and its full error, and the relevant slice of `graph.json` with anything sensitive removed. `graph.json` holds no cell values, so it is generally safe to share, but filenames and column headers can still be revealing, so read it before you paste it.
