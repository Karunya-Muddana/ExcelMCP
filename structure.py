import difflib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from excelmcp.graph_client import (
    close_session,
    create_session,
    get_sheet_names,
    get_used_range,
    list_excel_files,
)
from excelmcp.storage import atomic_write_text, get_config_dir, log, read_json


def get_graph_path() -> Path:
    return get_config_dir() / "graph.json"


def load_graph() -> dict[str, Any]:
    """Loads the structure graph, degrading to empty if the file is corrupt.

    A truncated graph.json used to raise JSONDecodeError out of every single
    tool call, leaving the server permanently broken with an opaque error.
    """
    graph = read_json(get_graph_path(), {"workspaces": {}})
    if not isinstance(graph, dict) or not isinstance(graph.get("workspaces"), dict):
        return {"workspaces": {}}
    return graph


def save_graph(graph: dict[str, Any]) -> None:
    atomic_write_text(get_graph_path(), json.dumps(graph, indent=2))


def normalise_name(name: Any) -> str:
    """Case- and whitespace-insensitive form of a sheet or column name.

    'Sales ', 'sales' and 'SALES' all normalise to 'sales'. Used to detect the
    near-miss naming that makes cross-file operations silently partial.
    """
    return " ".join(str(name).split()).casefold()


def fuzzy_name_candidates(
    target: str, names: Iterable[str], limit: int = 3
) -> list[str]:
    """Names that plausibly mean the same thing as target, best match first.

    Catches three kinds of drift: pure case/whitespace variants ('sales '),
    containment ('Sales 2024' for 'Sales'), and small typos ('Slaes'). These
    are *suggestions only* — callers must never fold a fuzzy match into a
    result as if it were the real thing.
    """
    target_n = normalise_name(target)
    if not target_n:
        return []
    scored: list[tuple[float, str]] = []
    for name in names:
        name_n = normalise_name(name)
        ratio = difflib.SequenceMatcher(None, target_n, name_n).ratio()
        if (
            name_n == target_n
            or target_n in name_n
            or name_n in target_n
            or ratio >= 0.75
        ):
            scored.append((ratio, name))
    scored.sort(key=lambda pair: -pair[0])
    return [name for _, name in scored[:limit]]


def sheet_name_variants(files: dict[str, Any]) -> dict[str, dict[str, list[str]]]:
    """Groups of sheet names across a workspace that normalise to the same string.

    Returns {normalised: {raw_spelling: [files using it]}} for every group with
    more than one raw spelling — the fragmentation an agent should see *before*
    it runs a cross-file aggregate.
    """
    groups: dict[str, dict[str, list[str]]] = {}
    for file_name, file_info in files.items():
        for sheet_name in file_info.get("sheets", {}):
            groups.setdefault(normalise_name(sheet_name), {}).setdefault(
                sheet_name, []
            ).append(file_name)
    return {norm: variants for norm, variants in groups.items() if len(variants) > 1}


def _deduplicate_columns(columns: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            result.append(col)
    return result


def detect_header_row(used_range_data: dict[str, Any]) -> tuple[int, list[str]]:
    """
    Detects the first row that looks like a header: >= 2 non-empty cells where
    at least half are strings. Skips single-cell title rows.
    """
    values = used_range_data.get("values", [])
    if not values:
        return 0, []

    for idx, row in enumerate(values):
        non_empty = [c for c in row if c is not None and str(c).strip()]
        string_cells = [c for c in non_empty if isinstance(c, str)]

        if len(non_empty) >= 2 and len(string_cells) >= len(non_empty) / 2:
            columns = [
                str(c).strip() if c is not None and str(c).strip() else f"Column_{i}"
                for i, c in enumerate(row)
            ]
            return idx, _deduplicate_columns(columns)

    first_row = values[0]
    columns = [
        str(c).strip() if c is not None and str(c).strip() else f"Column_{i}"
        for i, c in enumerate(first_row)
    ]
    return 0, _deduplicate_columns(columns)


def generate_sheet_description(file_name: str, sheet_name: str, columns: list[str]) -> str:
    cols_str = ", ".join(columns)
    return f"Data sheet '{sheet_name}' in workbook '{file_name}'. Contains columns: {cols_str}."


async def discover_structure(folder_path: str) -> dict[str, Any]:
    """
    Scans the folder, reads sheet headers, and builds the structure graph.
    Only column names and header positions are persisted — never cell values.
    """
    folder_path = folder_path.rstrip("/")
    graph = load_graph()
    graph.setdefault("workspaces", {})

    files = await list_excel_files(folder_path)
    log(f"[Scan] Discovered {len(files)} files in {folder_path}. Analyzing structure...")

    # Built fresh rather than merged into the previous scan, so workbooks deleted
    # from OneDrive stop appearing in the graph as phantom files.
    discovered: dict[str, Any] = {}

    for f in files:
        file_name = f["name"]
        item_id = f["id"]

        entry: dict[str, Any] = {
            "item_id": item_id,
            "last_scanned": datetime.now(timezone.utc).isoformat(),
            "sheets": {},
        }

        # A workbook session makes the per-sheet reads consistent and faster, but
        # it is an optimisation: if it cannot be created (file locked, unsupported
        # format) fall back to sessionless reads instead of aborting the scan.
        session_id = ""
        try:
            session_id = await create_session(item_id)
        except Exception as exc:
            log(f"[Scan] No workbook session for '{file_name}' ({exc}) — reading directly.")

        try:
            sheets = await get_sheet_names(item_id, session_id or None)
            for sheet_name in sheets:
                try:
                    data = await get_used_range(item_id, sheet_name, session_id or None)
                    header_idx, columns = detect_header_row(data)
                    if not columns:
                        log(f"[Scan] Sheet '{sheet_name}' in '{file_name}' is empty — skipped.")
                        continue
                    values = data.get("values", [])
                    entry["sheets"][sheet_name] = {
                        "header_row": header_idx + 1,
                        "columns": columns,
                        # A count, not cell values — recorded so inspect_file
                        # can report size without a live call. Labelled
                        # as_of_last_scan wherever it is surfaced.
                        "approx_row_count": max(0, len(values) - header_idx - 1),
                        "description": generate_sheet_description(
                            file_name, sheet_name, columns
                        ),
                    }
                except Exception as e:
                    log(f"[Scan] Skipping sheet '{sheet_name}' in '{file_name}': {e}")
        except Exception as e:
            log(f"[Scan] Skipping file '{file_name}': {e}")
            continue
        finally:
            if session_id:
                try:
                    await close_session(item_id, session_id)
                except Exception:
                    pass

        discovered[file_name] = entry

    workspace = {
        "files": discovered,
        "relationships": graph["workspaces"].get(folder_path, {}).get("relationships", []),
    }
    graph["workspaces"][folder_path] = workspace

    save_graph(graph)
    log(f"[Scan] Structure graph saved to {get_graph_path()}")
    return workspace
