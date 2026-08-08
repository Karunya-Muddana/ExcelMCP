"""The ExcelMCP command-line experience.

`excelmcp-setup` with no arguments runs the full guided wizard: authenticate,
choose a OneDrive folder, index it, and register the server with every AI agent
found on the machine. Subcommands expose the individual steps for scripting.

Everything is rendered with `rich`. Progress and prompts go to stderr-safe
console objects so this module can never interfere with the JSON-RPC stdio
stream used by the server itself.
"""

from __future__ import annotations

import argparse
import asyncio
import re
import sys
from typing import Optional

from rich.align import Align
from rich.console import Console, Group
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.rule import Rule
from rich.syntax import Syntax
from rich.table import Table
from rich.text import Text

from excelmcp import __version__
from excelmcp.agents import (
    AgentSpec,
    build_registry,
    install,
    manual_snippet,
    resolve_command,
    uninstall,
)

def _supports_unicode() -> bool:
    """Whether stdout can encode the box-drawing and status glyphs.

    Windows consoles still default to cp1252, where printing the banner raises
    UnicodeEncodeError and crashes the wizard before it does anything. We try to
    switch the stream to UTF-8 first, then verify by actually encoding a sample.
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
    except Exception:
        pass

    encoding = (getattr(sys.stdout, "encoding", "") or "ascii").lower()
    if "utf" in encoding:
        return True
    try:
        "█✓●○▸·✗".encode(encoding)
        return True
    except (UnicodeEncodeError, LookupError):
        return False


UNICODE = _supports_unicode()
console = Console()

_BANNER_UNICODE = r"""
 ███████╗██╗  ██╗ ██████╗███████╗██╗     ███╗   ███╗ ██████╗██████╗
 ██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║     ████╗ ████║██╔════╝██╔══██╗
 █████╗   ╚███╔╝ ██║     █████╗  ██║     ██╔████╔██║██║     ██████╔╝
 ██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║     ██║╚██╔╝██║██║     ██╔═══╝
 ███████╗██╔╝ ██╗╚██████╗███████╗███████╗██║ ╚═╝ ██║╚██████╗██║
 ╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝╚═╝     ╚═╝ ╚═════╝╚═╝
"""

_BANNER_ASCII = r"""
  _____                 _ __  __  ____ ____
 | ____|_  _____ ___   | |  \/  |/ ___|  _ \
 |  _| \ \/ / __/ _ \  | | |\/| | |   | |_) |
 | |___ >  < (_|  __/  | | |  | | |___|  __/
 |_____/_/\_\___\___|  |_|_|  |_|\____|_|
"""

BANNER = _BANNER_UNICODE if UNICODE else _BANNER_ASCII

# Every glyph has an ASCII fallback so the CLI stays readable on a legacy console.
G = {
    "ok": "✓" if UNICODE else "+",
    "warn": "!",
    "fail": "✗" if UNICODE else "x",
    "dot": "·" if UNICODE else "-",
    "registered": "●" if UNICODE else "*",
    "detected": "○" if UNICODE else "o",
    "selected": "▸" if UNICODE else ">",
}

ACTION_STYLE = {
    "added": ("[bold green]installed[/]", G["ok"]),
    "updated": ("[bold green]updated[/]", G["ok"]),
    "would-add": ("[cyan]would install[/]", G["dot"]),
    "would-update": ("[cyan]would update[/]", G["dot"]),
    "removed": ("[bold yellow]removed[/]", G["ok"]),
    "absent": ("[dim]not present[/]", G["dot"]),
    "skipped": ("[bold yellow]skipped[/]", G["warn"]),
    "failed": ("[bold red]failed[/]", G["fail"]),
}


# ----------------------------------------------------------------- chrome


def header() -> None:
    console.print()
    console.print(Align.center(Text(BANNER, style="bold cyan")))
    console.print(
        Align.center(
            Text.assemble(
                ("Live Excel intelligence for AI agents", "bold white"),
                ("  ·  ", "dim"),
                (f"v{__version__}", "dim"),
            )
        )
    )
    console.print()


def step(number: int, total: int, title: str) -> None:
    console.print()
    console.print(
        Rule(
            Text.assemble(
                (f" Step {number}/{total} ", "bold black on cyan"),
                ("  ", ""),
                (title, "bold white"),
            ),
            style="cyan",
            align="left",
        )
    )
    console.print()


def ok(msg: str) -> None:
    console.print(f"  [bold green]{G['ok']}[/] {msg}")


def warn(msg: str) -> None:
    console.print(f"  [bold yellow]{G['warn']}[/] {msg}")


def fail(msg: str) -> None:
    console.print(f"  [bold red]{G['fail']}[/] {msg}")


def info(msg: str) -> None:
    console.print(f"  [dim]{G['dot']}[/] [dim]{msg}[/]")


# ------------------------------------------------------------- agent table


def agent_table(specs: list[AgentSpec], selected: Optional[set[str]] = None) -> Table:
    table = Table(
        show_header=True,
        header_style="bold cyan",
        box=None,
        padding=(0, 2),
        expand=False,
    )
    table.add_column("", width=1)
    table.add_column("Agent", style="bold white")
    table.add_column("Status")
    table.add_column("Config", style="dim")

    for spec in specs:
        installed = spec.is_installed()
        configured = spec.is_configured()

        if configured:
            status = "[green]already registered[/]"
            mark = f"[green]{G['registered']}[/]"
        elif installed:
            status = "[cyan]detected[/]"
            mark = f"[cyan]{G['detected']}[/]"
        else:
            status = "[dim]not found[/]"
            mark = f"[dim]{G['dot']}[/]"

        if selected is not None and spec.key in selected:
            mark = f"[bold green]{G['selected']}[/]"

        path = str(spec.config)
        home = str(spec.config.home())
        if path.startswith(home):
            path = "~" + path[len(home):]

        table.add_row(mark, spec.label, status, path)
    return table


def results_table(rows: list[tuple[AgentSpec, dict]]) -> Table:
    table = Table(show_header=True, header_style="bold cyan", box=None, padding=(0, 2))
    table.add_column("", width=1)
    table.add_column("Agent", style="bold white")
    table.add_column("Result")
    table.add_column("Note", style="dim")

    for spec, res in rows:
        label, glyph = ACTION_STYLE.get(res["action"], ("[dim]?[/]", "·"))
        note = ""
        if res.get("backup"):
            note = f"backup: {res['backup'].name}"
        if not res["ok"]:
            note = res.get("detail", "")
        table.add_row(glyph, spec.label, label, note)
    return table


# -------------------------------------------------------------- selection


def choose_agents(specs: list[AgentSpec], assume_yes: bool) -> list[AgentSpec]:
    detected = [s for s in specs if s.is_installed()]

    if not detected:
        warn("No supported AI agents detected on this machine.")
        info("You can still register ExcelMCP by hand — snippets are printed below.")
        return []

    console.print(agent_table(specs))
    console.print()

    if assume_yes:
        ok(f"Registering with all {len(detected)} detected agent(s).")
        return detected

    console.print(
        "  [bold white]Which agents should ExcelMCP be registered with?[/]"
    )
    console.print(
        "  [dim]Enter to accept all detected, or a comma-separated list of names.[/]"
    )
    console.print(f"  [dim]Available: {', '.join(s.key for s in detected)}[/]")
    console.print()

    raw = Prompt.ask("  [cyan]agents[/]", default="all", show_default=True).strip()
    if raw.lower() in ("all", ""):
        return detected

    wanted = {p.strip().lower() for p in raw.split(",") if p.strip()}
    by_key = {s.key: s for s in detected}
    chosen = [by_key[k] for k in wanted if k in by_key]
    unknown = wanted - set(by_key)
    for u in unknown:
        warn(f"Unknown or undetected agent '{u}' — skipped.")
    return chosen


def show_manual(specs: list[AgentSpec], folder: str) -> None:
    if not specs:
        return
    console.print()
    console.print(
        Panel(
            "These agents were not configured automatically. "
            "Paste the snippet into the file shown.",
            title="[bold yellow]Manual setup[/]",
            border_style="yellow",
        )
    )
    for spec in specs:
        console.print()
        console.print(f"  [bold white]{spec.label}[/] [dim]→ {spec.config}[/]")
        lexer = {"json": "json", "yaml": "yaml", "toml": "toml"}[spec.fmt]
        console.print(
            Panel(
                Syntax(manual_snippet(spec, folder), lexer, theme="ansi_dark"),
                border_style="dim",
            )
        )


# ------------------------------------------------------------------ steps


async def do_auth() -> bool:
    from excelmcp.auth import get_token

    status = console.status("[cyan]Waiting for Microsoft sign-in...", spinner="dots")

    def announce(flow: dict) -> None:
        """Shows a device code. Called again with a new code if the first one is lost.

        The spinner is a live display that owns the bottom of the terminal, so it
        has to be stopped before anything else writes or the code can be painted
        over and never seen.
        """
        status.stop()
        uri = flow.get("verification_uri", "https://microsoft.com/devicelogin")
        console.print()
        console.print(
            Panel(
                f"Open [bold cyan]{uri}[/]\n"
                f"Enter code  [bold black on bright_cyan] {flow['user_code']} [/]",
                title="[bold]Sign in to Microsoft[/]",
                border_style="cyan",
                padding=(1, 3),
            )
        )
        status.start()

    try:
        status.start()
        await get_token(on_device_code=announce)
    except Exception as exc:
        status.stop()
        fail(str(exc))
        return False
    status.stop()
    ok("Signed in to Microsoft.")
    return True


def normalize_folder(raw: str) -> str:
    """Converts user input into a OneDrive-relative path.

    The most common setup mistake is pasting the *local* synced folder
    (`C:\\Users\\me\\OneDrive\\ERP` or `/Users/me/OneDrive/ERP`) instead of the
    path relative to the drive root. Graph resolves paths server-side, so a local
    path 404s with a message that does not explain the real problem. We detect a
    local path, and if it contains a OneDrive segment we recover the remainder.
    """
    folder = raw.strip().strip('"').strip("'")
    if not folder:
        return "/"

    looks_local = bool(re.match(r"^[A-Za-z]:[\\/]", folder)) or folder.startswith("\\\\")
    if looks_local or "onedrive" in folder.lower():
        parts = re.split(r"[\\/]+", folder)
        for i, part in enumerate(parts):
            if "onedrive" in part.lower():
                remainder = [p for p in parts[i + 1:] if p]
                recovered = "/" + "/".join(remainder)
                warn(f"That looks like a local path. Using [bold]{recovered}[/] instead.")
                info("ExcelMCP addresses folders relative to your OneDrive root.")
                folder = recovered
                break
        else:
            if looks_local:
                warn("That looks like a local filesystem path, not a OneDrive path.")
                info("Use a path relative to your OneDrive root, e.g. /ERP")

    folder = folder.replace("\\", "/")
    if not folder.startswith("/"):
        folder = "/" + folder
    while "//" in folder:
        folder = folder.replace("//", "/")
    return folder.rstrip("/") or "/"


def prompt_folder(preset: Optional[str]) -> str:
    if preset:
        return normalize_folder(preset)
    console.print("  [dim]The OneDrive folder holding your .xlsx files.[/]")
    console.print("  [dim]Examples: /ERP   /Finance/2026   /  (entire drive)[/]")
    console.print()
    return normalize_folder(Prompt.ask("  [cyan]folder[/]", default="/"))


async def do_scan(folder: str) -> bool:
    from excelmcp.embeddings import collection_count, update_embeddings
    from excelmcp.structure import discover_structure

    info("First run downloads the embedding model (~130 MB); this is one-time.")
    console.print()
    try:
        with console.status(
            f"[cyan]Scanning [white]{folder}[/] on OneDrive...", spinner="dots"
        ):
            workspace = await discover_structure(folder)
    except Exception as exc:
        fail(f"Scan failed: {exc}")
        return False

    files = workspace.get("files", {})
    if not files:
        fail(f"No .xlsx files found in '{folder}'.")
        info("Check the path, and that the files are synced to OneDrive.")
        return False

    sheets = sum(len(f["sheets"]) for f in files.values())
    ok(f"Indexed [bold]{len(files)}[/] workbook(s), [bold]{sheets}[/] sheet(s).")

    # A sheet holding several table bodies is one where a plain total adds up
    # blocks that were never meant to be summed. Say so at scan time, while the
    # user is looking, rather than leaving it for an agent to trip over.
    multi_region = [
        (file_name, sheet_name, len(sheet_info.get("regions", [])))
        for file_name, file_info in files.items()
        for sheet_name, sheet_info in file_info.get("sheets", {}).items()
        if len(sheet_info.get("regions", [])) > 1
    ]
    if multi_region:
        info(
            f"{len(multi_region)} sheet(s) hold more than one table. "
            f"Totalling one of these whole will sum unrelated blocks together:"
        )
        for file_name, sheet_name, count in multi_region[:10]:
            info(f"  {file_name} / {sheet_name} — {count} table regions")
        if len(multi_region) > 10:
            info(f"  ...and {len(multi_region) - 10} more.")

    try:
        with console.status("[cyan]Building semantic index...", spinner="dots"):
            update_embeddings(folder, files)
    except Exception as exc:
        fail(f"Embedding failed: {exc}")
        return False

    count = collection_count(folder)
    if count == 0:
        fail("Embedding validation failed — index is empty after scan.")
        return False
    ok(f"Semantic index ready ({count} sheet descriptions).")
    return True


def do_install(
    specs: list[AgentSpec], folder: str, dry_run: bool
) -> list[tuple[AgentSpec, dict]]:
    rows: list[tuple[AgentSpec, dict]] = []
    for spec in specs:
        res = install(spec, folder, dry_run=dry_run)
        rows.append((spec, res))
    return rows


def finish(folder: str, rows: list[tuple[AgentSpec, dict]]) -> None:
    succeeded = [s for s, r in rows if r["ok"] and r["action"] in ("added", "updated")]

    console.print()
    console.print(Rule(style="green"))
    console.print()
    console.print(
        Align.center(Text("ExcelMCP is ready", style="bold green"))
    )
    console.print()

    if succeeded:
        hints = Group(
            *[
                Text.assemble((f"  {s.label}  ", "bold white"), (s.restart_hint, "dim"))
                for s in succeeded
            ]
        )
        console.print(
            Panel(hints, title="[bold]Restart these to load the tools[/]", border_style="green")
        )

    examples = Group(
        Text("  Ask your agent things like:", style="dim"),
        Text(""),
        Text('  "What files do I have in this workspace?"', style="white"),
        Text('  "Total quantity by material across all files"', style="white"),
        Text('  "Which products have stock below 100?"', style="white"),
    )
    console.print(Panel(examples, title="[bold]Try it[/]", border_style="cyan"))

    console.print(
        f"  [dim]Workspace:[/] [white]{folder}[/]   "
        f"[dim]Diagnose anytime:[/] [white]excelmcp-setup doctor[/]"
    )
    console.print()


# --------------------------------------------------------------- commands


async def cmd_wizard(args: argparse.Namespace) -> int:
    header()
    specs = build_registry()
    total = 4

    step(1, total, "Microsoft authentication")
    if not await do_auth():
        return 1

    step(2, total, "Choose your OneDrive workspace")
    folder = prompt_folder(args.folder)
    ok(f"Workspace set to [bold]{folder}[/]")

    step(3, total, "Scan and index")
    if not await do_scan(folder):
        return 1

    step(4, total, "Connect your AI agents")
    if args.only:
        wanted = {k.strip().lower() for k in args.only.split(",")}
        chosen = [s for s in specs if s.key in wanted]
        missing = wanted - {s.key for s in specs}
        for m in missing:
            warn(f"Unknown agent '{m}'.")
    else:
        chosen = choose_agents(specs, args.yes)

    rows = do_install(chosen, folder, args.dry_run)
    if rows:
        console.print()
        console.print(results_table(rows))

    show_manual([s for s, r in rows if not r["ok"]], folder)
    finish(folder, rows)
    return 0


async def cmd_install(args: argparse.Namespace) -> int:
    header()
    specs = build_registry()
    folder = prompt_folder(args.folder) if not args.folder else args.folder

    if args.only:
        wanted = {k.strip().lower() for k in args.only.split(",")}
        chosen = [s for s in specs if s.key in wanted]
    else:
        chosen = choose_agents(specs, args.yes)

    rows = do_install(chosen, folder, args.dry_run)
    console.print()
    console.print(results_table(rows))
    show_manual([s for s, r in rows if not r["ok"]], folder)
    return 0


async def cmd_uninstall(args: argparse.Namespace) -> int:
    header()
    specs = build_registry()
    targets = [s for s in specs if s.is_configured()]
    if not targets:
        ok("ExcelMCP is not registered with any agent.")
        return 0

    console.print(agent_table(targets))
    console.print()
    if not args.yes and not Confirm.ask("  [cyan]Remove ExcelMCP from these?[/]"):
        info("Cancelled.")
        return 0

    rows = [(s, uninstall(s)) for s in targets]
    console.print()
    console.print(results_table(rows))
    return 0


async def cmd_doctor(args: argparse.Namespace) -> int:
    header()
    console.print(Rule("Diagnostics", style="cyan", align="left"))
    console.print()

    command, cmd_args = resolve_command()
    rendered = " ".join([command, *cmd_args])
    if shutil_which_ok(command):
        ok(f"Server command resolves: [white]{rendered}[/]")
    else:
        fail(f"Server command not found: {rendered}")

    from excelmcp.auth import get_token_cache_path
    from excelmcp.structure import get_graph_path, load_graph

    if get_token_cache_path().exists():
        ok("Microsoft token cache present.")
    else:
        warn("Not authenticated yet — run `excelmcp-setup`.")

    graph = load_graph()
    workspaces = graph.get("workspaces", {})
    if workspaces:
        ok(f"Structure graph: {len(workspaces)} workspace(s) at {get_graph_path()}")
        for name, ws in workspaces.items():
            files = ws.get("files", {})
            sheets = sum(len(f.get("sheets", {})) for f in files.values())
            info(f"{name} — {len(files)} file(s), {sheets} sheet(s)")
    else:
        warn("No workspace indexed — run `excelmcp-setup`.")

    from excelmcp.embeddings import METADATA_PATH, VECTORS_PATH

    if VECTORS_PATH.exists() and METADATA_PATH.exists():
        ok("Semantic index present.")
    else:
        warn("Semantic index missing — run `excelmcp-setup`.")

    console.print()
    console.print(Rule("Agent registration", style="cyan", align="left"))
    console.print()
    console.print(agent_table(build_registry()))
    console.print()
    return 0


def shutil_which_ok(command: str) -> bool:
    from pathlib import Path

    return Path(command).exists()


async def cmd_list(args: argparse.Namespace) -> int:
    header()
    console.print(agent_table(build_registry()))
    console.print()
    console.print(
        f"  [dim]{G['registered']} registered   "
        f"{G['detected']} detected, not registered   "
        f"{G['dot']} not installed[/]"
    )
    console.print()
    return 0


# ------------------------------------------------------------------ entry


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excelmcp-setup",
        description="Set up ExcelMCP and connect it to your AI agents.",
    )
    parser.add_argument("--version", action="version", version=f"excelmcp {__version__}")
    sub = parser.add_subparsers(dest="command")

    def common(p: argparse.ArgumentParser) -> None:
        p.add_argument("--folder", help="OneDrive folder, e.g. /ERP")
        p.add_argument("--only", help="Comma-separated agent keys, e.g. hermes,claude-code")
        p.add_argument("--yes", "-y", action="store_true", help="No prompts")
        p.add_argument("--dry-run", action="store_true", help="Show changes, write nothing")

    common(parser)
    parser.set_defaults(func=cmd_wizard)

    p_install = sub.add_parser("install", help="Register with agents only (no rescan)")
    common(p_install)
    p_install.set_defaults(func=cmd_install)

    p_uninstall = sub.add_parser("uninstall", help="Remove ExcelMCP from agent configs")
    p_uninstall.add_argument("--yes", "-y", action="store_true")
    p_uninstall.set_defaults(func=cmd_uninstall)

    p_doctor = sub.add_parser("doctor", help="Diagnose the installation")
    p_doctor.set_defaults(func=cmd_doctor)

    p_list = sub.add_parser("list-agents", help="Show detected agents")
    p_list.set_defaults(func=cmd_list)

    return parser


async def run_setup() -> None:
    """Runs the full wizard with defaults. Kept for `excelmcp.wizard:run_setup`."""
    await cmd_wizard(build_parser().parse_args([]))


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    for attr, default in (
        ("folder", None),
        ("only", None),
        ("yes", False),
        ("dry_run", False),
    ):
        if not hasattr(args, attr):
            setattr(args, attr, default)

    try:
        raise SystemExit(asyncio.run(args.func(args)))
    except KeyboardInterrupt:
        console.print()
        console.print("  [yellow]Cancelled.[/]")
        raise SystemExit(130)


if __name__ == "__main__":
    main()
