# Routines

Scheduled and repeatable jobs. Each file in this folder is one routine: what it is for, the prompt, how to schedule it, and what usually goes wrong with it.

| Routine | Cadence | What it does |
|---|---|---|
| [daily-inventory-check.md](daily-inventory-check.md) | every weekday morning | Flags anything low, out of stock, or newly at risk. |
| [weekly-sales-digest.md](weekly-sales-digest.md) | Monday morning | Week over week performance by region and product. |
| [month-end-reconciliation.md](month-end-reconciliation.md) | first working day | Cross-file totals with a per-file breakdown that must agree. |
| [data-quality-audit.md](data-quality-audit.md) | weekly, cheap | Structural problems in the workbooks themselves. |

## How to run them

### Claude Code scheduled agents

If your Claude Code build has the scheduling skill, this is the least effort path. It runs the prompt on a cron schedule in the cloud and reports back.

```
/schedule
```

Then give it the cadence and paste the routine prompt. Scheduled runs are non-interactive, so the prompt must be fully self-contained: name the folder, name the files, and say what the output should look like. A prompt that would prompt you a follow-up question will just stall.

### Cron plus a script

The portable option. Drive the server directly with an MCP client, as shown in [hosts.md](../guides/hosts.md), and put the result wherever you want it.

```bash
# 07:30 on weekdays
30 7 * * 1-5 /path/to/venv/bin/python /path/to/inventory_check.py >> /var/log/excelmcp.log 2>&1
```

The auth token lives in `~/.excelmcp/token.json` and belongs to a specific user account, so run the job as that user. A cron job running as root will not find it.

### Task Scheduler on Windows

```powershell
$action  = New-ScheduledTaskAction -Execute "C:\path\to\venv\Scripts\python.exe" `
                                   -Argument "C:\path\to\inventory_check.py"
$trigger = New-ScheduledTaskTrigger -Daily -At 7:30am
Register-ScheduledTask -TaskName "ExcelMCP Inventory Check" -Action $action -Trigger $trigger
```

Run it as your own user, for the same token reason, and tick the option to run whether logged on or not.

### GitHub Actions

Workable but awkward. The runner has no `~/.excelmcp`, so the token cache has to come from a secret, and refresh tokens expire, so the secret needs rotating. Fine for a job you can afford to have break occasionally. Not what you want for month-end close.

## Writing your own

The routines here follow a shape that has proven reliable for unattended runs.

**Pin the folder explicitly.** Never rely on `EXCELMCP_DEFAULT_FOLDER` in a scheduled job. If the env is not what you assumed, an explicit folder fails loudly instead of quietly reporting on the wrong workspace.

**Orient before fetching.** Start with `get_workspace_graph`. It costs nothing and it means the routine survives a column being renamed, rather than failing on a hardcoded string.

**Say what to do when data is missing.** An unattended agent with an unanswerable question will improvise. Tell it explicitly to report the gap instead. This is the single highest value line in any scheduled prompt.

**Demand the timestamp.** Every output should carry `fetched_at`. A report with no timestamp is indistinguishable from a report that silently failed and reused yesterday's numbers.

**Fix the output format.** Specify markdown, or a table, or JSON. Non-interactive output usually gets piped somewhere, and a chatty preamble breaks the thing downstream.

**Bound the work.** Say how many rows, how many items, how long. Without a bound, a routine that normally returns five rows will one day return four hundred.

## A note on scope

Everything here reads. There is no write path in ExcelMCP: it requests `Files.Read.All` and nothing else, so no routine can modify a workbook, and a misconfigured schedule cannot corrupt your data. The worst case is a wrong number in a report, which argues for the verification step that each routine below includes.
