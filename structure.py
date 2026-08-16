import asyncio
import difflib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Optional

from excelmcp.graph_client import (
    close_session,
    create_session,
    get_range,
    get_sheet_names,
    get_used_range,
    list_excel_files,
)
from excelmcp.ranges import build_range, parse_range
from excelmcp.regions import derive_sheet_regions, extract_sheet_references
from excelmcp.storage import atomic_write_text, get_config_dir, log, read_json


# Bumped whenever a scan starts recording something readers depend on. v3 adds
# per-sheet `regions`: without them every sheet is treated as one table, which
# is how a multi-section sheet silently returns a six-fold overcount. A graph
# written by an older version is not upgradable in place — the regions come
# from formulas nobody fetched — so readers refuse it and name the fix.
GRAPH_SCHEMA_VERSION = 3


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


def require_current_schema(workspace: dict[str, Any], folder_path: str) -> None:
    """Raises when a workspace was scanned by a version that knew less than this one.

    Following the precedent set for v0.2's range-address fields: an out-of-date
    graph fails loudly with the one command that fixes it, rather than quietly
    answering from a model of the workbook that no longer holds.
    """
    version = workspace.get("schema_version", 0)
    if not isinstance(version, int) or version < GRAPH_SCHEMA_VERSION:
        raise ValueError(
            f"The structure graph for '{folder_path}' predates v0.3 "
            f"(schema {version}, expected {GRAPH_SCHEMA_VERSION}) and has no "
            f"region map, so multi-section sheets would be read as one table. "
            f"Run scan_workspace on '{folder_path}' once to rebuild it."
        )


def save_graph(graph: dict[str, Any]) -> None:
    atomic_write_text(get_graph_path(), json.dumps(graph, indent=2))


# Chemical and unit names arrive both ways in one workbook: the sheet is called
# 'H2SO4 & Soda Ash' while a cell inside it says 'Sulphuric Acid (H₂SO₄)'. Folded
# to their ASCII digits, a query saying H2SO4 reaches both.
_DIGIT_FOLD = str.maketrans(
    {
        **{chr(0x2080 + d): str(d) for d in range(10)},  # ₀..₉
        "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
        "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
    }
)


def normalise_name(name: Any) -> str:
    """Case-, whitespace- and subscript-insensitive form of a sheet or column name.

    'Sales ', 'sales' and 'SALES' all normalise to 'sales'; 'H₂SO₄' and 'H2SO4'
    both become 'h2so4'. Used to detect the near-miss naming that makes
    cross-file operations silently partial.
    """
    return " ".join(str(name).translate(_DIGIT_FOLD).split()).casefold()


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
    """Turns one raw sheet row into cleaned, deduplicated column names.

    Internal whitespace is collapsed, not just trimmed: a header cell wrapped
    across two lines in Excel arrives as 'Qty\\n(Kgs)', and a column an agent
    cannot type is a column it cannot query. The exact spelling stays available
    via :func:`raw_columns_from_row`.
    """
    columns = [
        " ".join(str(c).split()) if c is not None and str(c).strip() else f"Column_{i}"
        for i, c in enumerate(row)
    ]
    return _deduplicate_columns(columns)


def raw_columns_from_row(row: list[Any]) -> list[str]:
    """Header names exactly as the sheet spells them, positionally aligned with
    :func:`columns_from_row` and deduplicated the same way.

    Kept so that writing back to a cell, or matching a name a user copied out of
    Excel, can use the real string rather than the display-friendly one.
    """
    columns = [
        str(c) if c is not None and str(c).strip() else f"Column_{i}"
        for i, c in enumerate(row)
    ]
    return _deduplicate_columns(columns)


# ------------------------------------------------------- header detection
#
# Header rows are scored, not first-matched. The old rule — first row with >= 2
# non-empty cells, half of them strings — fails in both directions on real
# workbooks: it takes a summary block near the top as the header on sheets that
# open with one, and it has no way to argue with a formula-derived candidate
# that is a row late whenever an `Opening Balance` line sits between the header
# and the first row the totals cover.
#
# Scoring every plausible row against the data beneath it settles both cases
# without privileging either source.

_DATE_TEXT = re.compile(r"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}|^\d{4}-\d{2}-\d{2}")

# A header cell is a label. Past this length it is a sentence, and a sentence
# in the top rows is a note to the reader.
_MAX_LABEL_LEN = 60

# Rows below a candidate read as its body when scoring: enough to establish
# each column's type profile, few enough that a totals block does not dilute it.
_BODY_SAMPLE = 20

# How deep a header can plausibly sit. Banners, section titles and summary
# blocks stack up, but a header below this is not a header.
_MAX_HEADER_DEPTH = 40

# A candidate must be beaten by this margin to be displaced, so an established
# answer does not flip on noise.
_OVERRIDE_MARGIN = 0.08

HEADER_CONFIDENT = 0.55
HEADER_WEAK = 0.35


def cell_kind(value: Any) -> str:
    """empty | number | date | text — the shape of one cell, not its content."""
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "text"
    if isinstance(value, (int, float)):
        return "number"
    text = str(value).strip()
    if not text:
        return "empty"
    # Graph hands numbers back as strings often enough that treating '14000'
    # as text would let a label/value row pass as a header.
    try:
        float(text.replace(",", "").replace("\u20b9", "").replace("%", ""))
        return "number"
    except ValueError:
        pass
    if _DATE_TEXT.match(text):
        return "date"
    return "text"


def score_header_row(row: list[Any], body: list[list[Any]]) -> float:
    """How much this row behaves like the header of ``body``. 0.0 to ~1.0.

    Five signals, none decisive alone:

    fill        A header names most of the columns its table uses.
    text_ratio  Headers are labels; data rows carry numbers and dates.
    uniqueness  Two columns rarely share a name; two data cells often share
                a value.
    contrast    What actually separates a header from the data row below it:
                in the columns where this row holds text, the body should hold
                something other than text. 'Qty (Kgs)' over a column of numbers
                is a header. 'Opening Balance' over a column of dates is one
                half of a label/value pair.
    penalties   Numeric cells and sentence-length strings both count against.
    """
    width = max(len(row), 1)
    kinds = [cell_kind(c) for c in row]
    non_empty = [k for k in kinds if k != "empty"]
    if len(non_empty) < 2:
        # A single-cell row is a title or a section banner, never a header.
        return 0.0

    texts = [str(c).strip() for c, k in zip(row, kinds) if k == "text"]
    fill = len(non_empty) / width
    text_ratio = len(texts) / len(non_empty)
    uniqueness = len({t.casefold() for t in texts}) / max(len(texts), 1)
    numeric_penalty = sum(
        1 for k in non_empty if k in ("number", "date")
    ) / len(non_empty)
    long_penalty = sum(1 for t in texts if len(t) > _MAX_LABEL_LEN) / max(
        len(texts), 1
    )

    contrast = 0.0
    text_cols = [i for i, k in enumerate(kinds) if k == "text"]
    if body and text_cols:
        differing = counted = 0
        for i in text_cols:
            below = [cell_kind(r[i]) for r in body if i < len(r)]
            filled = [k for k in below if k != "empty"]
            if not filled:
                continue
            counted += 1
            if sum(1 for k in filled if k != "text") / len(filled) >= 0.5:
                differing += 1
        if counted:
            contrast = differing / counted

    score = (
        0.30 * fill
        + 0.30 * text_ratio
        + 0.15 * uniqueness
        + 0.25 * contrast
        - 0.35 * numeric_penalty
        - 0.10 * long_penalty
    )
    return max(0.0, score)


def resolve_header_row(
    values: list[list[Any]],
    candidates: Optional[list[int]] = None,
) -> tuple[int, float, str]:
    """Best header row index, its score, and a confidence label.

    ``candidates`` are indices another source already proposed — the
    formula-derived row, a previous heuristic's pick. They are scored like any
    other row and given a small margin, so an established answer is not
    displaced by a coin-flip.
    """
    if not values:
        return 0, 0.0, "unreliable"

    candidates = candidates or []
    depth = min(len(values), _MAX_HEADER_DEPTH)
    considered = set(range(depth)) | {c for c in candidates if 0 <= c < len(values)}

    scores = {
        i: score_header_row(values[i], values[i + 1 : i + 1 + _BODY_SAMPLE])
        for i in considered
    }
    # Earlier rows win ties: a header sits above its data, and a repeated
    # header band further down should not outrank the first one.
    best_idx = max(scores, key=lambda i: (scores[i], -i))
    best_score = scores[best_idx]

    for c in candidates:
        if c in scores and scores[c] + _OVERRIDE_MARGIN >= best_score:
            best_idx, best_score = c, scores[c]
            break

    if best_score >= HEADER_CONFIDENT:
        confidence = "confident"
    elif best_score >= HEADER_WEAK:
        confidence = "weak"
    else:
        confidence = "unreliable"
    return best_idx, round(best_score, 3), confidence


def placeholder_ratio(columns: list[str]) -> float:
    """Share of column names that are positional placeholders.

    A high ratio means the detected row named almost nothing, and any tool
    asked to aggregate over those names is working from a layout nobody
    resolved — worth surfacing rather than letting a confident number out.
    """
    if not columns:
        return 1.0
    placeholders = sum(1 for c in columns if re.fullmatch(r"Column_\d+(_\d+)?", c))
    return round(placeholders / len(columns), 3)


def detect_header_row(used_range_data: dict[str, Any]) -> tuple[int, list[str], bool]:
    """Detects the header row by scoring every plausible candidate.

    Returns (row_index, columns, found). found=False means nothing scored well
    enough to rely on — callers scanning a bounded window use it to decide
    whether a wider read is worth it, and the caller records the confidence.
    """
    values = used_range_data.get("values", [])
    if not values:
        return 0, [], False

    idx, score, confidence = resolve_header_row(values)
    return idx, columns_from_row(values[idx]), confidence != "unreliable"


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
                    "kind": "join_key",
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


def _columns_at(sheet_info: dict[str, Any], address: str) -> list[str]:
    """Column names a formula's target address lands on, best-effort.

    Empty when the target sheet has no cached range address, or the address
    falls outside the columns recorded for it.
    """
    target_address = sheet_info.get("used_range_address")
    columns = sheet_info.get("columns") or []
    if not target_address or not columns:
        return []
    try:
        base_col = parse_range(target_address)[0]
        col1, _, col2, _ = parse_range(address)
    except ValueError:
        return []
    names = []
    for col in range(col1, col2 + 1):
        idx = col - base_col
        if 0 <= idx < len(columns):
            names.append(columns[idx])
    return names


def formula_relationships(files: dict[str, Any]) -> list[dict[str, Any]]:
    """Cross-sheet dependencies the workbook states outright in its formulas.

    ``='CMC Statement'!C48`` on the Dashboard is not an inference — it is the
    author wiring two sheets together, and it carries confidence 1.0 for that
    reason. It is *not* a join key, though: it says "this cell reads that cell",
    not "these two columns hold the same entities". Both live in the same
    ``relationships`` list so an agent can see the whole dependency picture, and
    they are told apart by ``kind`` — ``"reference"`` here against the implicit
    ``"join_key"`` of :func:`infer_relationships`. Anything choosing join keys
    must filter on it; joining on a reference edge would be nonsense.
    """
    found: list[dict[str, Any]] = []
    for file_name, file_info in files.items():
        sheets = file_info.get("sheets", {})
        by_normalised = {normalise_name(name): name for name in sheets}
        for sheet_name, sheet_info in sheets.items():
            for reference in sheet_info.get("sheet_references", []) or []:
                raw_target = reference.get("sheet", "")
                target = (
                    raw_target
                    if raw_target in sheets
                    else by_normalised.get(normalise_name(raw_target), "")
                )
                if not target or target == sheet_name:
                    continue
                addresses = reference.get("addresses", [])
                columns: list[str] = []
                for address in addresses:
                    for name in _columns_at(sheets[target], address):
                        if name not in columns:
                            columns.append(name)
                found.append(
                    {
                        "left": {
                            "file": file_name,
                            "sheet": sheet_name,
                            "column": None,
                        },
                        "right": {
                            "file": file_name,
                            "sheet": target,
                            "column": columns[0] if len(columns) == 1 else None,
                        },
                        "kind": "reference",
                        "confidence": 1.0,
                        "source": "formula",
                        "evidence": {
                            "addresses": addresses[:10],
                            "reference_count": len(addresses),
                            "target_columns": columns[:10],
                        },
                    }
                )
    found.sort(key=lambda r: (r["left"]["file"], r["left"]["sheet"], r["right"]["sheet"]))
    return found[:_MAX_RELATIONSHIPS]


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

    # Formulas for the whole sheet, not just the window. They compress far
    # better than values — a column of them is the same expression repeated per
    # row — and they are the only free evidence of where each table body ends.
    formulas: list[list[Any]] = []
    try:
        formula_data = await get_used_range(
            item_id, sheet_name, session_id, select="formulas"
        )
        formulas = formula_data.get("formulas") or []
    except Exception as exc:
        # A sheet whose formulas cannot be read still has a usable header; it
        # just falls back to the single-region heuristic.
        log(f"[Scan] No formulas for '{sheet_name}' ({exc}) — regions unavailable.")

    regions, unclaimed = derive_sheet_regions(formulas, row1, row2)
    references = extract_sheet_references(formulas)

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
    header_source = "heuristic"
    block = window
    block_row1 = row1

    # The formula-derived body start and the scored candidate are both just
    # candidates. body_start - 1 is right whenever the header sits directly
    # above the totalled block, and a row late whenever an Opening Balance line
    # sits between them — so it is scored against the data like anything else
    # rather than trusted outright.
    header_score = 0.0
    header_confidence = "weak"
    primary_header = regions[0]["header_row"] if regions else None
    if primary_header is not None:
        if not (row1 <= primary_header <= min(row2, row1 + _HEADER_WINDOW_ROWS - 1)):
            # The formula-derived candidate sits outside the detection window,
            # so read a small block starting above it — far enough up to catch
            # a header the totals do not point directly at.
            block_row1 = max(row1, primary_header - 8)
            block = await get_range(
                item_id,
                sheet_name,
                build_range(
                    col1,
                    block_row1,
                    min(col2, col1 + _SCAN_WINDOW_COLS - 1),
                    min(row2, block_row1 + _HEADER_WINDOW_ROWS - 1),
                ),
                session_id,
                select="values,numberFormat",
            )
        block_values = block.get("values") or []
        candidate_idx = primary_header - block_row1
        if block_values:
            resolved, header_score, header_confidence = resolve_header_row(
                block_values, [candidate_idx] if candidate_idx >= 0 else None
            )
            header_idx = resolved
            columns = columns_from_row(block_values[resolved])
            found = header_confidence != "unreliable"
            header_source = (
                "formula" if resolved == candidate_idx else "formula+scored"
            )
        else:
            # The read came back empty. Keep the window's answer rather than
            # indexing into nothing.
            block, block_row1 = window, row1
    else:
        block_values = block.get("values") or []
        if block_values:
            _, header_score, header_confidence = resolve_header_row(
                block_values, [header_idx]
            )

    column_types: dict[str, str] = {}
    if found:
        # numberFormat is only fetched for the block; if the header lives
        # beyond it (rare), date detection is skipped rather than paying for
        # formats on a full-sheet read.
        column_types = detect_date_columns(
            block.get("values", []),
            block.get("numberFormat", []),
            header_idx,
            columns,
        )

    if not found and row_count > _HEADER_WINDOW_ROWS:
        full = await get_used_range(item_id, sheet_name, session_id)
        header_idx, columns, found = detect_header_row(full)

    if not columns:
        return None

    header_abs = block_row1 + header_idx  # absolute 1-based sheet row
    raw_columns = list(columns)

    if column_count > _SCAN_WINDOW_COLS:
        # The detection window is capped at 52 columns; re-read the header row
        # at full width so wide sheets don't lose column names.
        header_row_address = build_range(col1, header_abs, col2, header_abs)
        header_data = await get_range(
            item_id, sheet_name, header_row_address, session_id
        )
        header_values = (header_data.get("values") or [[]])[0]
        if header_values:
            columns = columns_from_row(header_values)
            raw_columns = raw_columns_from_row(header_values)
    else:
        block_values = block.get("values") or []
        if 0 <= header_idx < len(block_values):
            raw_columns = raw_columns_from_row(block_values[header_idx])

    sampled_values: dict[str, list[str]] = {}
    if found:
        # Sampling is confined to the first table body when one is known, so
        # section titles and totals rows below it never become "values this
        # column contains" in the routing description.
        sample_start = header_abs + 1
        sample_end = min(row2, sample_start + _SAMPLE_ROWS - 1)
        if regions:
            body_end = int(regions[0]["body"].split(":")[1])
            sample_end = min(sample_end, body_end)
        if sample_start <= sample_end:
            sample_address = build_range(
                col1,
                sample_start,
                min(col2, col1 + _SCAN_WINDOW_COLS - 1),
                sample_end,
            )
            sample = await get_range(
                item_id, sheet_name, sample_address, session_id, select="values"
            )
            sampled_values = sample_column_values(
                sample.get("values", []), columns
            )

    entry: dict[str, Any] = {
        # 1-based offset into the used range, not an absolute sheet row — the
        # convention every reader (query_engine, lookup) already relies on.
        # `regions` uses absolute sheet rows; the two differ by row1 - 1.
        "header_row": header_abs - row1 + 1,
        "header_source": header_source,
        "columns": columns,
        "used_range_address": address,
        "row_count": row_count,
        "column_count": column_count,
        # A count, not cell values — recorded so inspect_file can report
        # size without a live call. Labelled as_of_last_scan where surfaced.
        "approx_row_count": max(0, row2 - header_abs),
        "regions": regions,
        "unclaimed_rows": unclaimed,
        # Tier 1 guesses every header from a formula's body start; nothing has
        # confirmed it and nothing knows what the regions *mean*. An agent
        # should be able to see that a sheet has not been interpreted yet.
        "layout_confidence": "unconfirmed",
        # How much the detected header row actually behaved like a header, and
        # how much of the result is positional filler. A sheet whose columns are
        # mostly Column_N was not interpreted, and a tool asked to aggregate
        # over those names should say so instead of returning a number.
        "header_confidence": header_confidence,
        "header_score": header_score,
        "placeholder_ratio": placeholder_ratio(columns),
    }
    if raw_columns != columns:
        entry["columns_raw"] = raw_columns
    if references:
        entry["sheet_references"] = references
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
        "schema_version": GRAPH_SCHEMA_VERSION,
        "files": discovered,
        # Recomputed from scratch each scan; user-declared relationships live
        # in relationships.yaml and are merged in at read time, not persisted.
        # Formula-declared references come first: they are facts, not guesses.
        "relationships": (
            formula_relationships(discovered) + infer_relationships(discovered)
        ),
    }
    graph["workspaces"][folder_path] = workspace

    save_graph(graph)
    log(f"[Scan] Structure graph saved to {get_graph_path()}")
    return workspace