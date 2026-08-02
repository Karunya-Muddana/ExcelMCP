# ExcelMCP

Universal live Excel intelligence layer for AI agents. Connects to OneDrive via Microsoft Graph API, indexes workbook structure as semantic embeddings, and answers natural language queries against live data.

**Core invariant**: structure is cached (sheet names, column headers). Data is always fetched live. Zero cell values stored anywhere.

---

## Prerequisites

- Python 3.10+
- Microsoft 365 account with OneDrive
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

---

## Installation

```bash
# From the excelmcp/ directory
uv sync          # install dependencies
uv build         # build distributable wheel
pip install dist/excelmcp-0.1.0-*.whl
```

---

## Setup

Run the interactive setup wizard once:

```bash
excelmcp-setup
```

This will:
1. Open device-flow Microsoft authentication
2. Prompt for your OneDrive folder path (e.g. `/ERP`)
3. Scan all `.xlsx` files and build the structure graph
4. Auto-configure Hermes (`~/.hermes/config.yaml`) and Claude Code (`~/.claude.json`) if detected

---

## Configuration

| Environment variable | Default | Description |
|---|---|---|
| `EXCELMCP_CLIENT_ID` | built-in app | Azure AD application client ID |
| `EXCELMCP_TENANT_ID` | `common` | Azure AD tenant (use `common` for personal accounts) |

---

## Available Tools

| Tool | Description |
|---|---|
| `scan_workspace` | Crawl a folder, extract structure, build embeddings |
| `get_workspace_graph` | Return full structure graph (no live data) |
| `inspect_file` | Return sheet/column info for one file (no live data) |
| `query` | Natural language question — RAG-routed, live data |
| `filter_sheet` | Fetch sheet live and filter rows by conditions |
| `aggregate` | Fetch sheet live and run groupby sum/mean/count |
| `cross_file_aggregate` | Combine multiple sheets and aggregate |

---

## Example Usage (Claude Code)

```
scan_workspace("/ERP")

query("What are the top 10 products by sales value?", folder_path="/ERP")

filter_sheet("Inventory.xlsx", "Stock", {"Status": "Low"}, folder_path="/ERP", limit=100)

aggregate("Sales.xlsx", "Q1", group_by="Region", value_col="Revenue",
          operation="sum", folder_path="/ERP")

cross_file_aggregate(
    sources=[
        {"file": "Sales.xlsx", "sheet": "Q1"},
        {"file": "Sales.xlsx", "sheet": "Q2"}
    ],
    value_col="Revenue", operation="sum",
    group_by="Region", folder_path="/ERP"
)
```

---

## Data Architecture

```
~/.excelmcp/
  token.json       — MSAL token cache (auth only; written 0600)
  graph.json       — structure graph: file IDs, sheet names, column headers
  vectors.npy      — embeddings of sheet descriptions (NumPy cosine search)
  metadata.json    — sheet/file/workspace labels for each embedding
```

Cell values are never stored. Every query fetches fresh data from the Microsoft Graph API.

---

## Running Tests

```bash
# Offline unit tests — no network or credentials needed
pytest tests/test_unit.py

# Live integration tests against a real scanned workspace (opt-in)
EXCELMCP_TEST_FOLDER=/ERP pytest tests/test_live_integration.py -v
```

---

## License

MIT
