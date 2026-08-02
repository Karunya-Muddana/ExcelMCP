# Host Guide

Per-host setup, verification, and quirks. `excelmcp-setup` writes these config entries for you, so most people never need this file. Read it when auto-detection missed your host, when you want to know what was written to your config, or when the server is registered but not working.

## What gets written

Every host gets the same underlying thing: a command to run, arguments, and an environment variable holding your default folder. Only the surrounding shape differs.

The most common shape, used by Claude Desktop, Cursor, Windsurf, Gemini CLI, and Cline:

```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "excelmcp",
      "args": [],
      "env": { "EXCELMCP_DEFAULT_FOLDER": "/ERP" }
    }
  }
}
```

Claude Code and VS Code want an explicit transport:

```json
{
  "mcpServers": {
    "excelmcp": {
      "type": "stdio",
      "command": "excelmcp",
      "args": [],
      "env": { "EXCELMCP_DEFAULT_FOLDER": "/ERP" }
    }
  }
}
```

`EXCELMCP_DEFAULT_FOLDER` is what lets you say "what is low on stock" instead of passing `folder_path` on every single call. Without it, every tool call needs the folder spelled out.

### The command resolution problem

Agents launched from a GUI often do not inherit your shell's PATH, so a bare `excelmcp` resolves fine in a terminal and fails silently inside Claude Desktop or Cursor. The wizard handles this by writing an absolute interpreter path with `-m excelmcp` when it cannot guarantee the console script is reachable:

```json
{
  "command": "/absolute/path/to/python",
  "args": ["-m", "excelmcp"]
}
```

If you are writing config by hand and the server will not start, this is the first thing to try. `python -m excelmcp` always works when the package is installed into that interpreter.

---

## Per host

### Claude Code

Config: `~/.claude.json`, container `mcpServers`, explicit `stdio` type.

```bash
excelmcp-setup install --only claude-code
```

Verify with `/mcp` in a session. You should see ExcelMCP connected with seven tools.

Put workspace-specific guidance in `CLAUDE.md` at your project root. For heavier use, a dedicated subagent in `.claude/agents/` keeps spreadsheet work out of your main context:

```markdown
---
name: excel-analyst
description: Answers questions about the ERP spreadsheets in OneDrive /ERP.
tools: get_workspace_graph, inspect_file, query, filter_sheet, aggregate, cross_file_aggregate
---

<paste the trimmed prompt from ../system-prompt.md, plus your workspace specifics>
```

Note the omission of `scan_workspace` from that tool list. A read-only analyst subagent has no business re-crawling the workspace, and leaving the tool out is more reliable than telling it not to.

Restart hint: run `claude` again, or `/mcp` to reconnect.

### Claude Desktop

Config: `claude_desktop_config.json`, under `%APPDATA%\Claude` on Windows, `~/Library/Application Support/Claude` on macOS, `~/.config/Claude` on Linux.

This is the host most likely to hit the PATH problem, since it launches from the dock or Start menu. If tools do not appear, switch to the absolute interpreter form above.

Restart hint: quit completely and reopen. Closing the window is not enough on macOS.

### Cursor

Config: `~/.cursor/mcp.json`, container `mcpServers`.

Verify under Settings, MCP. Workspace guidance goes in `.cursor/rules/excelmcp.mdc`.

Restart hint: reload the window.

### Windsurf

Config: `~/.codeium/windsurf/mcp_config.json`, container `mcpServers`.

Verify in the Cascade panel under MCP. Rules go in `.windsurfrules`.

### Gemini CLI

Config: `~/.gemini/settings.json`, container `mcpServers`.

Verify with `/mcp` in a session. Guidance goes in `GEMINI.md`.

### Codex CLI

Config: `~/.codex/config.toml`, container `mcp_servers`. This is the only TOML host, so the entry looks different:

```toml
[mcp_servers.excelmcp]
command = "excelmcp"
args = []

[mcp_servers.excelmcp.env]
EXCELMCP_DEFAULT_FOLDER = "/ERP"
```

Guidance goes in `AGENTS.md` at the project root.

Restart hint: start a new `codex` session.

### VS Code with Copilot

Config: `mcp.json` in the VS Code user directory, container `servers` rather than `mcpServers`, with an explicit `stdio` type.

Restart hint: reload the window.

### Cline

Config: `cline_mcp_settings.json` inside the extension's global storage, under `saoudrizwan.claude-dev/settings`.

Restart hint: reload VS Code, then open the Cline MCP panel.

### Continue

Config: `~/.continue/config.yaml`, container `mcpServers`. YAML, not JSON, despite what older Continue documentation says.

### Goose

Config: `~/.config/goose/config.yaml`, container `extensions`. Goose uses its own extension shape rather than the standard MCP server entry, which the wizard handles.

Restart hint: start a new `goose` session.

### Zed

Config: `settings.json` under `%APPDATA%\Zed` on Windows or `~/.config/zed` elsewhere, container `context_servers`.

Restart hint: restart Zed.

### Hermes

Config: `~/.hermes/config.yaml`, container `mcp_servers`. The entry includes `enabled: true`, which is part of the documented Hermes schema and is not optional: without it the server can be parsed and then ignored.

Restart hint: start a new Hermes session.

---

## Hosts the wizard does not know

Ask it to print the block and paste it yourself:

```bash
excelmcp-setup install --dry-run
```

This shows the exact JSON, YAML, or TOML for every detected host without writing anything. For an unknown host, the standard `mcpServers` shape is almost always right.

Anything that speaks MCP over stdio can run this server. The command is `excelmcp`, or `python -m excelmcp` if PATH is uncertain, and it speaks JSON-RPC on stdin and stdout. Nothing else may be printed to stdout or the host's parser desyncs, which is why the server never logs there.

## Programmatic use

You do not need an interactive host at all. Any MCP client library can spawn the server directly. With the Python MCP SDK:

```python
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

params = StdioServerParameters(
    command="python",
    args=["-m", "excelmcp"],
    env={"EXCELMCP_DEFAULT_FOLDER": "/ERP"},
)

async with stdio_client(params) as (read, write):
    async with ClientSession(read, write) as session:
        await session.initialize()
        graph = await session.call_tool("get_workspace_graph", {})
        rows = await session.call_tool("filter_sheet", {
            "file_name": "Inventory.xlsx",
            "sheet": "Stock",
            "conditions": {"Status": "Low"},
        })
```

This is the path to take for the scheduled [routines](../routines/), for a Slack bot, or for anything that needs to run without a human in the loop.

## Managing configs

```bash
excelmcp-setup list-agents            # what is installed and what is configured
excelmcp-setup install --only cursor  # one host, skip the rescan
excelmcp-setup install --dry-run      # show every change, write nothing
excelmcp-setup uninstall              # remove from every config
```

Existing config files are backed up before modification, and a config that exists but cannot be parsed is left alone rather than overwritten. If you have hand-edited JSON with a trailing comma somewhere, the wizard will refuse to touch that file and tell you, which is the correct outcome.
