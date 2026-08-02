# ExcelMCP

**A live Excel intelligence layer for AI agents.** Point it at a OneDrive folder and your agent can ask questions about those spreadsheets in plain English, against the numbers that are in them right now.

[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/protocol-MCP-000000)](https://modelcontextprotocol.io/)
[![Built with FastMCP](https://img.shields.io/badge/built%20with-FastMCP-4B32C3)](https://github.com/jlowin/fastmcp)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-0078D4?logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/graph/)
[![Status](https://img.shields.io/badge/status-alpha-orange)](https://github.com/Karunya-Muddana/ExcelMCP)
[![PRs welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/Karunya-Muddana/ExcelMCP/pulls)

---

## The problem this solves

Most spreadsheet integrations work by copying your data somewhere else. They ingest the workbook, chunk it, embed the cell values, and store the whole thing in a vector database. From that moment on your agent is answering questions about a snapshot. Someone updates the inventory sheet at 9am and the agent is still quoting Tuesday's numbers.

ExcelMCP splits the problem in two.

**Structure gets cached.** Filenames, sheet names, column headers, where the header row starts. This changes rarely, it is cheap to store, and it is what the agent needs in order to know *what to ask for*.

**Data never gets cached.** Not once. Every tool call that returns a number goes out to the Microsoft Graph API and pulls the current used range. There is no data cache to go stale, no sync job to fall behind, and no cell value written to disk anywhere in this project.

Every response carries a `metadata.fetched_at` timestamp and an `is_cached: false` flag so the model can see, in band, that it is looking at fresh data.

---

## How it works

```
  your agent
      |
      |  MCP over stdio
      v
  ExcelMCP server
      |
      +--> graph.json      structure only, read instantly, no network
      |
      +--> vectors.npy     embedded sheet descriptions, for routing questions
      |
      +--> Microsoft Graph ------> OneDrive
                                   every cell value, fetched on demand
```

A natural language question gets embedded, matched against the sheet descriptions by cosine similarity, and routed to the sheets most likely to hold the answer. Those sheets, and only those, get fetched live. Filtering and aggregation then happen in pandas on the freshly fetched frame.

---

## Requirements

- Python 3.10 or newer
- A Microsoft 365 account with OneDrive
- [uv](https://docs.astral.sh/uv/), or plain pip if you prefer

---

## Install

From the repository root:

```bash
git clone https://github.com/Karunya-Muddana/ExcelMCP.git
cd ExcelMCP

uv sync      # install dependencies
uv build     # build the wheel
pip install dist/excelmcp-0.1.0-py3-none-any.whl
```

Or install straight from source without building:

```bash
pip install .
```

There is no compiler step and no native extension to build. Vector search runs on a NumPy cosine scan rather than hnswlib, specifically so that `pip install` works on a machine with no C++ toolchain.

---

## Setup

Run the wizard once:

```bash
excelmcp-setup
```

It walks through four things:

1. Microsoft device-flow sign in. You get a code, you paste it into the browser, the token cache lands in `~/.excelmcp/token.json` with `0600` permissions.
2. Which OneDrive folder to index, for example `/ERP`.
3. A scan of every `.xlsx` in that folder to build the structure graph and the embeddings.
4. Detection of the AI agents already installed on your machine, and a written config entry for the ones you pick.

### Agents it can configure automatically

| Agent | Config file |
|---|---|
| Claude Code | `~/.claude.json` |
| Claude Desktop | `claude_desktop_config.json` |
| Cursor | `~/.cursor/mcp.json` |
| Windsurf | `~/.codeium/windsurf/mcp_config.json` |
| Gemini CLI | `~/.gemini/settings.json` |
| Codex CLI | `~/.codex/config.toml` |
| VS Code (Copilot) | VS Code user `mcp.json` |
| Cline | extension `cline_mcp_settings.json` |
| Continue | `~/.continue/config.yaml` |
| Goose | `~/.config/goose/config.yaml` |
| Zed | `~/.config/zed/settings.json` |
| Hermes | `~/.hermes/config.yaml` |

Existing config files are backed up before they are touched. If your agent is not on the list, the wizard prints the exact JSON or TOML block to paste in yourself.

### Other wizard commands

```bash
excelmcp-setup list-agents           # show what was detected
excelmcp-setup install --only cursor # register with one agent, skip the rescan
excelmcp-setup doctor                # diagnose a broken install
excelmcp-setup uninstall             # remove ExcelMCP from every agent config
excelmcp-setup --folder /ERP --yes   # fully non-interactive
excelmcp-setup --dry-run             # print the changes, write nothing
```

---

## Tools exposed to the agent

| Tool | Network | What it does |
|---|---|---|
| `get_workspace_graph` | none | Full structure of the workspace: files, sheets, columns. Instant. |
| `inspect_file` | none | Same, narrowed to one file. Instant. |
| `scan_workspace` | heavy | Re-crawls OneDrive and rebuilds structure plus embeddings. |
| `query` | live | Natural language question, semantically routed to the right sheets. |
| `filter_sheet` | live | Fetch one sheet, return rows matching conditions. |
| `aggregate` | live | Fetch one sheet, group and reduce it. |
| `cross_file_aggregate` | live | Fetch matching sheets from every file in parallel, then total. |

The two structure tools are free and instant because they read the local graph. Everything marked live goes to the API on every single call.

---

## Usage

Once the server is registered, you mostly just talk to your agent normally. Under the hood it makes calls like these.

Orient first. The agent should always do this before guessing at a column name, since no two companies name things the same way:

```python
get_workspace_graph(folder_path="/ERP")
```

Ask a question without knowing where the answer lives:

```python
query("what are the top 10 products by sales value", folder_path="/ERP")
```

Filter a known sheet:

```python
filter_sheet(
    file_name="Inventory.xlsx",
    sheet="Stock",
    conditions={"Status": "Low", "Quantity": "<50"},
    folder_path="/ERP",
    sort_by="Quantity",
    limit=100,
)
```

Supported condition operators, all ANDed together:

| Form | Meaning |
|---|---|
| `{"Col": "value"}` | exact match |
| `{"Col": "~value"}` | contains, literal substring, not a regex |
| `{"Col": ">100"}` | greater than |
| `{"Col": ">=100"}` | greater or equal |
| `{"Col": "<100"}` | less than |
| `{"Col": "<=100"}` | less or equal |

A column name that does not exist raises an error rather than quietly returning zero rows, which is the failure mode that makes an agent confidently report the wrong thing.

Group and reduce inside one file:

```python
aggregate(
    file_name="Sales.xlsx",
    sheet="Q1",
    group_by="Region",
    value_col="Revenue",
    operation="sum",
    folder_path="/ERP",
)
```

Total the same sheet across every file in the workspace:

```python
cross_file_aggregate(
    sheet="Q1",
    value_col="Revenue",
    operation="sum",
    folder_path="/ERP",
    conditions={"Status": "Closed"},
)
```

`cross_file_aggregate` returns a per-file breakdown alongside the total, plus `skipped_files` and a `warning` when some file could not be read. That way a partial total is visibly partial instead of silently wrong.

---

## Guardrails built into the server

The server ships a set of operating rules in its MCP instructions, which the host model reads before it makes its first call. They exist because these are the specific ways an LLM gets spreadsheet questions wrong:

- Never assume a filename, sheet name, or column name. Discover it from the graph.
- Never add up cross-file numbers mentally. Call `cross_file_aggregate` and let the tool do it.
- Never reach for `openpyxl`, `pandas.read_excel`, or the local filesystem. The files are not on this machine.
- Never sum a quantity column in transaction-style data without filtering by transaction type first.
- Treat numeric dates as Excel serials, offset from 1899-12-30.
- Check the `truncated` and `total_matched` fields before claiming a result is complete.

---

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `EXCELMCP_CLIENT_ID` | built in | Azure AD application client ID |
| `EXCELMCP_TENANT_ID` | `common` | Tenant. Use `common` for personal accounts. |
| `EXCELMCP_DEFAULT_FOLDER` | unset | Folder to use when a tool call omits `folder_path`. The wizard writes this into your agent config. |

The built in client ID is a public client used for device-code flow. It carries no secret, it is visible in every auth request by design, and it is safe to have in this repository. Swap it for your own app registration if you want the consent screen to carry your organisation's name.

---

## What lands on disk

```
~/.excelmcp/
  token.json      MSAL token cache. Auth material only, written 0600.
  graph.json      Structure graph: item IDs, sheet names, column headers.
  vectors.npy     Embedded sheet descriptions for semantic routing.
  metadata.json   Labels tying each embedding back to a sheet.
```

No cell value appears in any of these files. If you want to verify that claim rather than take it on faith, `graph.json` is small and readable, so go look.

On Windows, `os.chmod` only toggles the read-only bit, so the `0600` mode is a best effort there and the real protection is the default per-user ACL on `%USERPROFILE%`. On macOS and Linux the mode is applied to the temp file before any content is written to it, so the token never briefly exists as world readable.

---

## Tests

```bash
# offline unit tests, no network and no credentials required
pytest tests/test_unit.py

# live integration tests against a workspace you have already scanned, opt in
EXCELMCP_TEST_FOLDER=/ERP pytest tests/test_live_integration.py -v
```

The integration suite skips itself when `EXCELMCP_TEST_FOLDER` is unset, so a plain `pytest` run stays offline.

---

## Project layout

```
auth.py           MSAL device flow, token cache, proactive refresh
graph_client.py   Graph API wrapper, 429 backoff, session reuse
structure.py      Structure discovery, writes graph.json
embeddings.py     FastEmbed vectors plus NumPy cosine search
query_engine.py   Semantic routing, live fetch, pandas operations
main.py           FastMCP tool definitions and server entry point
cli.py            Setup wizard, agent detection, config writing
agents.py         Per agent config formats and file locations
storage.py        Atomic writes and config directory handling
```

---

## Contributing

Issues and pull requests are welcome. If you are adding support for another agent, `agents.py` is the only file you should need to touch: add an `AgentSpec` with the config path, the entry shape, and a detection hint.

---

## License

MIT. See [LICENSE](LICENSE).
