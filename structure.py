import asyncio
import difflib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from excelmcp.graph_client import (
    close_session,
    create_session,
    get_range,
    get_sheet_names,
    get_used_range,
    list_excel_files,
)
from excelmcp.ranges import build_range, parse_range
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


def columns_from_row(row: list[Any]) -> list[str]:
    """Turns one raw sheet row into cleaned, deduplicated column names."""
    columns = [
        str(c).strip() if c is not None and str(c).strip() else f"Column_{i}"
        for i, c in enumerate(row)
    ]
    return _deduplicate_columns(columns)


def detect_header_row(used_range_data: dict[str, Any]) -> tuple[int, list[str], bool]:
    """
    Detects the first row that looks like a header: >= 2 non-empty cells where
    at least half are strings. Skips single-cell title rows.

    Returns (row_index, columns, found). found=False means nothing matched the
    heuristic and the first row was used as a fallback — callers scanning a
    bounded window use it to decide whether a wider read is worth it.
    """
    values = used_range_data.get("values", [])
    if not values:
        return 0, [], False

    for idx, row in enumerate(values):
        non_empty = [c for c in row if c is not None and str(c).strip()]
        string_cells = [c for c in non_empty if isinstance(c, str)]

        if len(non_empty) >= 2 and len(string_cells) >= len(non_empty) / 2:
            return idx, columns_from_row(row), True

    return 0, columns_from_row(values[0]), False


# ------------------------------------------------------------ relationships

# A relationship is only ever *inferred* from two signals agreeing: the
# column names normalise to the same token (Material_ID ~ MaterialID ~
# 'material id') AND the sampled value sets overlap. Name match alone is an
# assumption; value overlap alone is coincidence; together they are evidence.
_RELATIONSHIP_MIN_OVERLAP = 0.3
_MAX_RELATIONSHIPS = 200


def _column_token(name: Any) -> str:
    """'Material_ID', 'materialid' and 'Material ID' all become 'materialid'."""
    return re.sub(r"[^a-z0-9]", "", str(name).casefold())


def infer_relationships(files: dict[str, Any]) -> list[dict[str, Any]]:
    """Cross-sheet key relationships, inferred from names plus sampled values.

    Only sampled (low-cardinality) columns participate — inference never runs
    on columns it has no value evidence for. Each relationship carries its
    confidence (the value-overlap ratio) and the evidence behind it.
    """
    entries = []
    for file_name, file_info in files.items():
        for sheet_name, sheet_info in file_info.get("sheets", {}).items():
            for column, values in sheet_info.get("sampled_values", {}).items():
                token = _column_token(column)
                normalised = {normalise_name(v) for v in values}
                if token and len(normalised) >= 2:
                    entries.append(
                        (file_name, sheet_name, column, token, normalised)
                    )

    relationships = []
    for i in range(len(entries)):
        for j in range(i + 1, len(entries)):
            a, b = entries[i], entries[j]
            if a[3] != b[3] or (a[0], a[1]) == (b[0], b[1]):
                continue
            shared = a[4] & b[4]
            overlap = len(shared) / min(len(a[4]), len(b[4]))
            if overlap < _RELATIONSHIP_MIN_OVERLAP:
                continue
            relationships.append(
                {
                    "left": {"file": a[0], "sheet": a[1], "column": a[2]},
                    "right": {"file": b[0], "sheet": b[1], "column": b[2]},
                    "confidence": round(overlap, 3),
                    "source": "inferred",
                    "evidence": {
                        "name_token": a[3],
                        "value_overlap": round(overlap, 3),
                        "left_sample_size": len(a[4]),
                        "right_sample_size": len(b[4]),
                    },
                }
            )
    relationships.sort(key=lambda r: -r["confidence"])
    return relationships[:_MAX_RELATIONSHIPS]


def get_relationships_yaml_path() -> Path:
    return get_config_dir() / "relationships.yaml"


def load_declared_relationships() -> list[dict[str, Any]]:
    """User-declared relationships from ~/.excelmcp/relationships.yaml.

    Inference covers a well-named workspace; the last stretch needs a human to
    state it once. Declared entries carry confidence 1.0 and outrank anything
    inferred. Expected shape:

        relationships:
          - left:  {file: A.xlsx, sheet: Stock,  column: Material}
            right: {file: B.xlsx, sheet: Prices, column: Material_ID}
    """
    path = get_relationships_yaml_path()
    if not path.exists():
        return []
    import yaml

    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except (OSError, yaml.YAMLError) as exc:
        log(f"[Relationships] Ignoring unreadable {path.name}: {exc}")
        return []
    if not isinstance(data, dict):
        return []

    declared = []
    for item in data.get("relationships", []) or []:
        left, right = item.get("left"), item.get("right")
        if not (
            isinstance(left, dict)
            and isinstance(right, dict)
            and all(k in left and k in right for k in ("file", "sheet", "column"))
        ):
            log(f"[Relationships] Skipping malformed entry in {path.name}: {item!r}")
            continue
        declared.append(
            {
                "left": {k: str(left[k]) for k in ("file", "sheet", "column")},
                "right": {k: str(right[k]) for k in ("file", "sheet", "column")},
                "confidence": 1.0,
                "source": "declared",
            }
        )
    return declared


def workspace_relationships(workspace: dict[str, Any]) -> list[dict[str, Any]]:
    """Declared relationships first, then inferred ones from the last scan."""
    return load_declared_relationships() + list(workspace.get("relationships", []))


# Value sampling bounds. Sampled values are the one deliberate exception to
# the "no cell values on disk" guarantee: a handful of distinct labels per
# low-cardinality column, persisted so that twenty structurally identical
# workbooks are distinguishable at routing time. They are never served as
# data — every answer still comes from a live fetch.
_SAMPLE_ROWS = 500
_MAX_SAMPLED_CARDINALITY = 50
_MAX_SAMPLED_VALUE_LEN = 60
_DESCRIPTION_VALUES_PER_COLUMN = 15


def sample_column_values(
    rows: list[list[Any]], columns: list[str]
) -> dict[str, list[str]]:
    """Distinct string values per low-cardinality column, from sampled rows.

    Only textual values count — quantities and serials say nothing about which
    sheet is which. Columns whose distinct count exceeds the cardinality cap
    (free text, IDs) are dropped entirely.
    """
    sampled: dict[str, list[str]] = {}
    for col_idx, col_name in enumerate(columns):
        distinct: dict[str, None] = {}
        over_cap = False
        for row in rows:
            cell = row[col_idx] if col_idx < len(row) else None
            if not isinstance(cell, str):
                continue
            value = " ".join(cell.split())
            if not value or len(value) > _MAX_SAMPLED_VALUE_LEN:
                continue
            distinct[value] = None
            if len(distinct) > _MAX_SAMPLED_CARDINALITY:
                over_cap = True
                break
        if distinct and not over_cap:
            sampled[col_name] = sorted(distinct)
    return sampled


def generate_sheet_description(
    file_name: str,
    sheet_name: str,
    columns: list[str],
    sampled_values: dict[str, list[str]] | None = None,
) -> str:
    """Text embedded for routing. Sampled values are what make twenty
    workbooks with identical schemas distinguishable — without them, cosine
    similarity between their descriptions approaches 1.0 and routing is
    effectively random."""
    cols_str = ", ".join(columns)
    parts = [
        f"Data sheet '{sheet_name}' in workbook '{file_name}'. "
        f"Contains columns: {cols_str}."
    ]
    for col_name, values in (sampled_values or {}).items():
        shown = values[:_DESCRIPTION_VALUES_PER_COLUMN]
        parts.append(f"{col_name} values include: {', '.join(shown)}.")
    return " ".join(parts)


_QUOTED_LITERAL_RE = re.compile(r'"[^"]*"')
_BRACKET_SECTION_RE = re.compile(r"\[[^\]]*\]")


def is_date_number_format(fmt: Any) -> bool:
    """True when an Excel number format renders a date or time.

    Date/time formats use y/m/d/h/s tokens ('m/d/yyyy', 'dd-mmm-yy hh:mm');
    numeric formats use 0/#/? placeholders. Quoted literals and [] sections
    (colors, locales, elapsed-time) are stripped first so '"day" 0' or
    '[$-409]d-mmm' classify by their real tokens.
    """
    if not fmt or not isinstance(fmt, str):
        return False
    cleaned = _BRACKET_SECTION_RE.sub("", _QUOTED_LITERAL_RE.sub("", fmt))
    # A format can carry ';'-separated sections (positive;negative;zero;text).
    # 'd-mmm;@' is a date format with a text fallback — classify per section.
    for section in cleaned.split(";"):
        lowered = section.lower()
        if "general" in lowered or "@" in section:
            continue
        if any(placeholder in section for placeholder in ("0", "#", "?")):
            continue
        if any(token in lowered for token in ("y", "m", "d", "h", "s")):
            return True
    return False


def detect_date_columns(
    values: list[list[Any]],
    number_formats: list[list[Any]],
    header_idx: int,
    columns: list[str],
) -> dict[str, str]:
    """Maps column name -> 'date' for columns whose cells carry date formats.

    Looks at the first non-empty data cell below the header in each column.
    Excel stores dates as serial floats; the numberFormat is the only signal
    that distinguishes a date column from a plain numeric one.
    """
    column_types: dict[str, str] = {}
    data_rows = range(header_idx + 1, len(values))
    for col_idx, col_name in enumerate(columns):
        for row_idx in data_rows:
            row = values[row_idx] if row_idx < len(values) else []
            cell = row[col_idx] if col_idx < len(row) else None
            if cell is None or (isinstance(cell, str) and not cell.strip()):
                continue
            fmt_row = (
                number_formats[row_idx] if row_idx < len(number_formats) else []
            )
            fmt = fmt_row[col_idx] if col_idx < len(fmt_row) else None
            if is_date_number_format(fmt):
                column_types[col_name] = "date"
            break
    return column_types


# Header detection reads a bounded window rather than the whole sheet: ten
# rows is where real headers live (title rows above them are rare and short),
# and 52 columns (A..AZ) covers the window; wider sheets get their header row
# fetched separately at full width.
_HEADER_WINDOW_ROWS = 10
_SCAN_WINDOW_COLS = 52


async def _scan_sheet_structure(
    item_id: str, sheet_name: str, session_id: Any
) -> dict[str, Any] | None:
    """Reads just enough of a sheet to learn its structure. Returns None if empty.

    Scanning used to download every sheet in full purely to find the header
    row index. Now: one metadata call for the used-range address and
    dimensions, one bounded window read for header detection, and a full read
    only if no header shows up in the first _HEADER_WINDOW_ROWS rows.
    """
    meta = await get_used_range(
        item_id, sheet_name, session_id, select="address,rowCount,columnCount"
    )
    address = str(meta.get("address") or "")
    row_count = int(meta.get("rowCount") or 0)
    column_count = int(meta.get("columnCount") or 0)
    if not address or row_count <= 0 or column_count <= 0:
        return None

    col1, row1, col2, row2 = parse_range(address)
    window_address = build_range(
        col1,
        row1,
        min(col2, col1 + _SCAN_WINDOW_COLS - 1),
        min(row2, row1 + _HEADER_WINDOW_ROWS - 1),
    )
    window = await get_range(
        item_id,
        sheet_name,
        window_address,
        session_id,
        select="values,numberFormat",
    )
    header_idx, columns, found = detect_header_row(window)

    column_types: dict[str, str] = {}
    if found:
        # numberFormat is only fetched for the window; if the header lives
        # beyond it (rare), date detection is skipped rather than paying for
        # formats on a full-sheet read.
        column_types = detect_date_columns(
            window.get("values", []),
            window.get("numberFormat", []),
            header_idx,
            columns,
        )

    if not found and row_count > _HEADER_WINDOW_ROWS:
        full = await get_used_range(item_id, sheet_name, session_id)
        header_idx, columns, found = detect_header_row(full)

    if not columns:
        return None

    if column_count > _SCAN_WINDOW_COLS:
        # The detection window is capped at 52 columns; re-read the header row
        # at full width so wide sheets don't lose column names.
        header_row_address = build_range(
            col1, row1 + header_idx, col2, row1 + header_idx
        )
        header_data = await get_range(
            item_id, sheet_name, header_row_address, session_id
        )
        header_values = (header_data.get("values") or [[]])[0]
        if header_values:
            columns = columns_from_row(header_values)

    sampled_values: dict[str, list[str]] = {}
    if found:
        sample_start = row1 + header_idx + 1
        if sample_start <= row2:
            sample_address = build_range(
                col1,
                sample_start,
                min(col2, col1 + _SCAN_WINDOW_COLS - 1),
                min(row2, sample_start + _SAMPLE_ROWS - 1),
            )
            sample = await get_range(
                item_id, sheet_name, sample_address, session_id, select="values"
            )
            sampled_values = sample_column_values(
                sample.get("values", []), columns
            )

    entry: dict[str, Any] = {
        "header_row": header_idx + 1,
        "columns": columns,
        "used_range_address": address,
        "row_count": row_count,
        "column_count": column_count,
        # A count, not cell values — recorded so inspect_file can report
        # size without a live call. Labelled as_of_last_scan where surfaced.
        "approx_row_count": max(0, row_count - header_idx - 1),
    }
    if column_types:
        entry["column_types"] = column_types
    if sampled_values:
        entry["sampled_values"] = sampled_values
    return entry


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

            # Sheets are read concurrently; the Graph client's shared
            # semaphore keeps at most max_concurrency() requests in flight,
            # so a 100-sheet workbook no longer costs 100 serial round-trips.
            async def read_sheet(sheet_name: str) -> Any:
                try:
                    return await _scan_sheet_structure(
                        item_id, sheet_name, session_id or None
                    )
                except Exception as exc:
                    return exc

            results = await asyncio.gather(*[read_sheet(s) for s in sheets])
            for sheet_name, sheet_entry in zip(sheets, results):
                if isinstance(sheet_entry, Exception):
                    log(f"[Scan] Skipping sheet '{sheet_name}' in '{file_name}': {sheet_entry}")
                    continue
                if sheet_entry is None:
                    log(f"[Scan] Sheet '{sheet_name}' in '{file_name}' is empty — skipped.")
                    continue
                sheet_entry["description"] = generate_sheet_description(
                    file_name,
                    sheet_name,
                    sheet_entry["columns"],
                    sheet_entry.get("sampled_values"),
                )
                entry["sheets"][sheet_name] = sheet_entry
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
        # Recomputed from scratch each scan; user-declared relationships live
        # in relationships.yaml and are merged in at read time, not persisted.
        "relationships": infer_relationships(discovered),
    }
    graph["workspaces"][folder_path] = workspace

    save_graph(graph)
    log(f"[Scan] Structure graph saved to {get_graph_path()}")
    return workspace
