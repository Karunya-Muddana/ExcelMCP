"""Registry of AI agents ExcelMCP can register itself with, and the logic to edit
their config files safely.

Design rules, all of which exist because agent configs are user-owned files that
must never be damaged:

* **Detect before writing.** An agent is only offered if evidence of an install
  exists (its config directory is present). We never create a config tree for an
  agent the user does not have.
* **Never clobber.** A config that exists but does not parse is left completely
  alone and reported, rather than being overwritten with a fresh one.
* **Always back up.** Any file we modify is copied to a timestamped sibling first.
* **Merge, don't replace.** Only our own key is added; every other server and
  unrelated setting in the file is preserved.

Config formats and locations drift between agent releases. `excelmcp-setup doctor`
re-reads every file and reports what is actually registered, and the wizard prints
a copy-pasteable snippet for anything it could not write, so a stale path here
degrades to a manual step rather than a silent failure.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

SERVER_KEY = "excelmcp"

HOME = Path.home()
_APPDATA = Path(os.environ.get("APPDATA", HOME / "AppData" / "Roaming"))


def _platform_path(windows: Path, macos: Path, linux: Path) -> Path:
    if sys.platform == "win32":
        return windows
    if sys.platform == "darwin":
        return macos
    return linux


def resolve_command() -> tuple[str, list[str]]:
    """Returns the (command, args) an agent should spawn to start the server.

    Prefers the absolute path to the installed `excelmcp` script. GUI-launched
    agents such as Claude Desktop do not inherit the shell's PATH, so a bare
    "excelmcp" often fails there with a spawn ENOENT that looks like a bug in the
    server. Falls back to `<python> -m excelmcp`, which always resolves.
    """
    found = shutil.which(SERVER_KEY)
    if found:
        return str(Path(found).resolve()), []
    return str(Path(sys.executable).resolve()), ["-m", "excelmcp"]


# --------------------------------------------------------------- entry shapes


def _entry_standard(command: str, args: list[str], env: dict[str, str]) -> dict:
    """The `mcpServers` shape used by Claude, Cursor, Windsurf, Gemini, Cline."""
    return {"command": command, "args": args, "env": env}


def _entry_stdio(command: str, args: list[str], env: dict[str, str]) -> dict:
    """Same, with an explicit transport type (Claude Code, VS Code)."""
    return {"type": "stdio", "command": command, "args": args, "env": env}


def _entry_goose(command: str, args: list[str], env: dict[str, str]) -> dict:
    return {
        "enabled": True,
        "type": "stdio",
        "cmd": command,
        "args": args,
        "envs": env,
        "name": SERVER_KEY,
        "description": "Universal live Excel intelligence layer",
    }


def _entry_zed(command: str, args: list[str], env: dict[str, str]) -> dict:
    return {
        "source": "custom",
        "command": command,
        "args": args,
        "env": env,
    }


def _entry_hermes(command: str, args: list[str], env: dict[str, str]) -> dict:
    """Hermes `mcp_servers` entry.

    `enabled: true` is part of Hermes' documented schema — without it a server can
    be registered but never started, which looks like the tools simply not existing.
    """
    return {
        "command": command,
        "args": args,
        "enabled": True,
        "env": env,
        "description": "Universal live Excel intelligence layer",
    }


# ------------------------------------------------------------------- specs


@dataclass
class AgentSpec:
    key: str
    label: str
    config: Path
    fmt: str  # "json" | "yaml" | "toml"
    container: list[str]  # key path to the server map
    entry: Callable[[str, list[str], dict[str, str]], Any]
    detect: list[Path] = field(default_factory=list)
    restart_hint: str = "Restart the app to pick up the new server."

    def is_installed(self) -> bool:
        return any(p.exists() for p in (self.detect or [self.config.parent]))

    def is_configured(self) -> bool:
        data = read_config(self)
        if data is None:
            return False
        node = data
        for part in self.container:
            if not isinstance(node, dict) or part not in node:
                return False
            node = node[part]
        return isinstance(node, dict) and SERVER_KEY in node


def _vscode_user_dir() -> Path:
    return _platform_path(
        _APPDATA / "Code" / "User",
        HOME / "Library" / "Application Support" / "Code" / "User",
        HOME / ".config" / "Code" / "User",
    )


def build_registry() -> list[AgentSpec]:
    vscode_user = _vscode_user_dir()
    claude_desktop_dir = _platform_path(
        _APPDATA / "Claude",
        HOME / "Library" / "Application Support" / "Claude",
        HOME / ".config" / "Claude",
    )
    zed_dir = _platform_path(
        _APPDATA / "Zed",
        HOME / ".config" / "zed",
        HOME / ".config" / "zed",
    )
    cline_dir = (
        vscode_user / "globalStorage" / "saoudrizwan.claude-dev" / "settings"
    )

    return [
        AgentSpec(
            key="claude-code",
            label="Claude Code",
            config=HOME / ".claude.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_stdio,
            detect=[HOME / ".claude.json", HOME / ".claude"],
            restart_hint="Run `claude` again, or /mcp to verify.",
        ),
        AgentSpec(
            key="claude-desktop",
            label="Claude Desktop",
            config=claude_desktop_dir / "claude_desktop_config.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[claude_desktop_dir],
            restart_hint="Quit Claude Desktop completely and reopen it.",
        ),
        AgentSpec(
            key="cursor",
            label="Cursor",
            config=HOME / ".cursor" / "mcp.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[HOME / ".cursor"],
            restart_hint="Reload Cursor, then check Settings → MCP.",
        ),
        AgentSpec(
            key="windsurf",
            label="Windsurf",
            config=HOME / ".codeium" / "windsurf" / "mcp_config.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[HOME / ".codeium" / "windsurf", HOME / ".codeium"],
            restart_hint="Reload Windsurf, then check Cascade → MCP.",
        ),
        AgentSpec(
            key="gemini-cli",
            label="Gemini CLI",
            config=HOME / ".gemini" / "settings.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[HOME / ".gemini"],
            restart_hint="Start a new `gemini` session, then /mcp.",
        ),
        AgentSpec(
            key="codex",
            label="Codex CLI",
            config=HOME / ".codex" / "config.toml",
            fmt="toml",
            container=["mcp_servers"],
            entry=_entry_standard,
            detect=[HOME / ".codex"],
            restart_hint="Start a new `codex` session.",
        ),
        AgentSpec(
            key="vscode",
            label="VS Code (Copilot)",
            config=vscode_user / "mcp.json",
            fmt="json",
            container=["servers"],
            entry=_entry_stdio,
            detect=[vscode_user],
            restart_hint="Reload the VS Code window.",
        ),
        AgentSpec(
            key="cline",
            label="Cline",
            config=cline_dir / "cline_mcp_settings.json",
            fmt="json",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[cline_dir],
            restart_hint="Reload VS Code, then open the Cline MCP panel.",
        ),
        AgentSpec(
            key="continue",
            label="Continue",
            config=HOME / ".continue" / "config.yaml",
            fmt="yaml",
            container=["mcpServers"],
            entry=_entry_standard,
            detect=[HOME / ".continue"],
            restart_hint="Reload your editor.",
        ),
        AgentSpec(
            key="goose",
            label="Goose",
            config=HOME / ".config" / "goose" / "config.yaml",
            fmt="yaml",
            container=["extensions"],
            entry=_entry_goose,
            detect=[HOME / ".config" / "goose"],
            restart_hint="Start a new `goose` session.",
        ),
        AgentSpec(
            key="zed",
            label="Zed",
            config=zed_dir / "settings.json",
            fmt="json",
            container=["context_servers"],
            entry=_entry_zed,
            detect=[zed_dir],
            restart_hint="Restart Zed.",
        ),
        AgentSpec(
            key="hermes",
            label="Hermes",
            config=HOME / ".hermes" / "config.yaml",
            fmt="yaml",
            container=["mcp_servers"],
            entry=_entry_hermes,
            detect=[HOME / ".hermes"],
            restart_hint="Start a new Hermes session.",
        ),
    ]


# ------------------------------------------------------------- read / write


class ConfigUnreadable(Exception):
    """The config file exists but could not be parsed — we must not touch it."""


def read_config(spec: AgentSpec) -> Optional[dict]:
    """Parses an agent config, returning {} if absent or None if unparseable."""
    if not spec.config.exists():
        return {}
    try:
        raw = spec.config.read_text(encoding="utf-8")
    except OSError:
        return None
    if not raw.strip():
        return {}

    try:
        if spec.fmt == "json":
            data = json.loads(raw)
        elif spec.fmt == "yaml":
            import yaml

            data = yaml.safe_load(raw)
        else:  # toml
            data = _load_toml(raw)
    except Exception:
        return None

    if data is None:
        return {}
    return data if isinstance(data, dict) else None


def _load_toml(raw: str) -> dict:
    try:
        import tomllib  # Python 3.11+
    except ModuleNotFoundError:
        import tomli as tomllib  # type: ignore[no-redef]
    return tomllib.loads(raw)


def backup(path: Path) -> Optional[Path]:
    if not path.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = path.with_name(f"{path.name}.excelmcp-backup-{stamp}")
    shutil.copy2(path, dest)
    return dest


def _atomic_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.excelmcp-tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)


def render_entry(spec: AgentSpec, folder: str) -> Any:
    command, args = resolve_command()
    env = {"EXCELMCP_DEFAULT_FOLDER": folder} if folder else {}
    return spec.entry(command, args, env)


def manual_snippet(spec: AgentSpec, folder: str) -> str:
    """The exact text a user would paste if automatic registration failed."""
    entry = render_entry(spec, folder)
    if spec.fmt == "toml":
        return _toml_block(entry)
    payload: Any = {SERVER_KEY: entry}
    for part in reversed(spec.container):
        payload = {part: payload}
    if spec.fmt == "yaml":
        import yaml

        return yaml.safe_dump(payload, sort_keys=False).rstrip()
    return json.dumps(payload, indent=2)


def _toml_block(entry: dict) -> str:
    command = entry.get("command", SERVER_KEY)
    args = entry.get("args", [])
    env = entry.get("env", {})
    lines = [
        f"[mcp_servers.{SERVER_KEY}]",
        f"command = {json.dumps(command)}",
        f"args = {json.dumps(args)}",
    ]
    if env:
        lines.append(f"[mcp_servers.{SERVER_KEY}.env]")
        for k, v in env.items():
            lines.append(f"{k} = {json.dumps(v)}")
    return "\n".join(lines)


def install(spec: AgentSpec, folder: str, dry_run: bool = False) -> dict[str, Any]:
    """Registers ExcelMCP in one agent's config.

    Returns a result dict with `ok`, `action`, `backup`, and `detail`. Raises
    nothing — every failure is reported so the wizard can show a manual snippet.
    """
    data = read_config(spec)
    if data is None:
        return {
            "ok": False,
            "action": "skipped",
            "backup": None,
            "detail": (
                f"{spec.config} exists but could not be parsed. "
                f"Left untouched — add the snippet below by hand."
            ),
        }

    already = spec.is_configured()

    if dry_run:
        return {
            "ok": True,
            "action": "would-update" if already else "would-add",
            "backup": None,
            "detail": str(spec.config),
        }

    try:
        if spec.fmt == "toml":
            return _install_toml(spec, folder, already)
        return _install_structured(spec, folder, data, already)
    except OSError as exc:
        return {"ok": False, "action": "failed", "backup": None, "detail": str(exc)}


def _install_structured(
    spec: AgentSpec, folder: str, data: dict, already: bool
) -> dict[str, Any]:
    node = data
    for part in spec.container:
        nxt = node.get(part)
        if nxt is None:
            nxt = {}
            node[part] = nxt
        elif not isinstance(nxt, dict):
            return {
                "ok": False,
                "action": "skipped",
                "backup": None,
                "detail": f"'{'.'.join(spec.container)}' is not an object in {spec.config}.",
            }
        node = nxt

    node[SERVER_KEY] = render_entry(spec, folder)

    backup_path = backup(spec.config)
    if spec.fmt == "yaml":
        import yaml

        text = yaml.safe_dump(data, sort_keys=False, allow_unicode=True)
    else:
        text = json.dumps(data, indent=2) + "\n"
    _atomic_write(spec.config, text)

    return {
        "ok": True,
        "action": "updated" if already else "added",
        "backup": backup_path,
        "detail": str(spec.config),
    }


def _install_toml(spec: AgentSpec, folder: str, already: bool) -> dict[str, Any]:
    """Appends a TOML table for our server.

    A new `[table]` header at end-of-file is always valid TOML regardless of what
    precedes it, so appending avoids depending on a TOML *writer* (the stdlib
    only reads TOML). If our table is already present we rewrite the file with
    the old block stripped, so re-running setup stays idempotent.
    """
    existing = spec.config.read_text(encoding="utf-8") if spec.config.exists() else ""
    block = _toml_block(render_entry(spec, folder))

    if already:
        kept: list[str] = []
        skipping = False
        for line in existing.splitlines():
            stripped = line.strip()
            if stripped.startswith("["):
                skipping = stripped.startswith(f"[mcp_servers.{SERVER_KEY}")
            if not skipping:
                kept.append(line)
        existing = "\n".join(kept).rstrip()

    text = (existing.rstrip() + "\n\n" + block + "\n") if existing.strip() else block + "\n"
    backup_path = backup(spec.config)
    _atomic_write(spec.config, text)
    return {
        "ok": True,
        "action": "updated" if already else "added",
        "backup": backup_path,
        "detail": str(spec.config),
    }


def uninstall(spec: AgentSpec) -> dict[str, Any]:
    """Removes only our own key, leaving the rest of the config intact."""
    if not spec.config.exists():
        return {"ok": True, "action": "absent", "backup": None, "detail": ""}

    if spec.fmt == "toml":
        existing = spec.config.read_text(encoding="utf-8")
        if f"[mcp_servers.{SERVER_KEY}" not in existing:
            return {"ok": True, "action": "absent", "backup": None, "detail": ""}
        kept: list[str] = []
        skipping = False
        for line in existing.splitlines():
            stripped = line.strip()
            if stripped.startswith("["):
                skipping = stripped.startswith(f"[mcp_servers.{SERVER_KEY}")
            if not skipping:
                kept.append(line)
        backup_path = backup(spec.config)
        _atomic_write(spec.config, "\n".join(kept).rstrip() + "\n")
        return {"ok": True, "action": "removed", "backup": backup_path, "detail": ""}

    data = read_config(spec)
    if data is None:
        return {"ok": False, "action": "skipped", "backup": None, "detail": "unparseable"}

    node = data
    for part in spec.container:
        if not isinstance(node, dict) or part not in node:
            return {"ok": True, "action": "absent", "backup": None, "detail": ""}
        node = node[part]
    if not isinstance(node, dict) or SERVER_KEY not in node:
        return {"ok": True, "action": "absent", "backup": None, "detail": ""}

    del node[SERVER_KEY]
    backup_path = backup(spec.config)
    if spec.fmt == "yaml":
        import yaml

        text = yaml.safe_dump(data, sort_keys=False, allow_unicode=True)
    else:
        text = json.dumps(data, indent=2) + "\n"
    _atomic_write(spec.config, text)
    return {"ok": True, "action": "removed", "backup": backup_path, "detail": ""}
