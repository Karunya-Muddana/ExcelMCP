# Agent Playbook

Everything in this folder is about the layer above the server: how to prompt an agent that has ExcelMCP attached, how to wire it into each host, and what to automate once it works.

The server itself is documented in the [root README](../README.md). This folder assumes it is already installed and answering.

## What is here

| File | Use it when |
|---|---|
| [system-prompt.md](system-prompt.md) | You are building a custom agent, subagent, or assistant and need a drop-in system prompt. |
| [prompts.md](prompts.md) | You want a copy-paste prompt for a specific job, sorted by what you are trying to get done. |
| [guides/getting-started.md](guides/getting-started.md) | It is your first session and you want to confirm the whole chain works. |
| [guides/hosts.md](guides/hosts.md) | You need per-host setup, verification, and quirks for Claude Code, Cursor, Codex, Hermes, and the rest. |
| [guides/query-patterns.md](guides/query-patterns.md) | You know what you want but not how to phrase it, or your results look wrong. |
| [guides/troubleshooting.md](guides/troubleshooting.md) | Something is broken and you want the error message decoded. |
| [routines/](routines/) | You want this to run on a schedule instead of on demand. |

## The one thing to understand first

ExcelMCP gives an agent two different kinds of tool, and almost every mistake comes from confusing them.

**Structure tools are free.** `get_workspace_graph` and `inspect_file` read a local JSON file. No network, no latency, no cost. An agent should call these constantly and never apologise for it.

**Data tools are live.** `query`, `filter_sheet`, `aggregate`, and `cross_file_aggregate` each hit the Microsoft Graph API at the moment they are called. They are slower, they are rate limited, and they are the only source of truth for any actual number.

A well behaved agent looks up the structure for free, then makes one precise live call. A badly behaved agent guesses a column name, gets an error, guesses again, and burns four API round trips discovering something that was sitting in a local file the whole time.

Everything in this folder is downstream of that idea.

## Naming note

There is an `agents.py` at the repository root. That is the module holding per-host config formats for the setup wizard. It is unrelated to this folder, and Python resolves the module ahead of this directory, so the two do not collide.
