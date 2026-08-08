"""Acceptance tests for Part D: single-cell lookup across ~100 sheets in one call.

The workspace here mirrors the worked example: 20 contract workbooks with an
identical five-sheet schema (100 sheets), each Rate Card holding 1000 rows by
30 columns. The Graph layer is faked at lookup.get_range with full request and
byte accounting, so these tests pin the two properties that make the feature
worth having: one tool call in, one provenanced cell out, at a tiny fraction
of the bytes a whole-sheet fetch moves.

Every failure mode must be explicit: a wrong number from this tool looks
maximally authoritative, so no response may ever be a bare value.
"""

import asyncio
import json

import pytest

from excelmcp import lookup, ranges
from excelmcp.structure import GRAPH_SCHEMA_VERSION as SCHEMA

N_ROWS = 1000  # data rows per Rate Card
N_COLS = 30

CLIENTS = [
    "BESTEX", "GREENFIELD", "VEDA", "ACME", "NORDWERK",
    "SUNCHEM", "POLYFAB", "OXIDANE", "KROMA", "LITHOS",
    "MERIDIAN", "QUANTA", "ZENITH", "HELIX", "ORBIT",
    "PINNACLE", "STRATUM", "VECTOR", "WYVERN", "AXIOM",
]

# file index -> {0-based data row -> (material, rate)}
SPECIALS = {
    0: {346: ("Titanium Dioxide", 142.5)},
    1: {100: ("Calcium Carbonate", 88.0)},
    2: {110: ("Calcium Carbonate", 88.0)},
    3: {120: ("Calcium Carbonate", 88.0)},
    4: {200: ("Zinc Oxide", 10.0)},
    5: {210: ("Zinc Oxide", 20.0)},
    6: {220: ("Zinc Oxide", 30.0)},
    7: {50: ("Sodium Chloride", 5.0), 60: ("Sodium Chloride", 7.0)},
}


def _file_name(idx):
    return f"CON{idx:03d}_{CLIENTS[idx]}.xlsx"


def _rate_card_grid(idx):
    client = CLIENTS[idx]
    header = ["Material", "Unit", "Contracted Rate", "Client"] + [
        f"Fill_{i}" for i in range(N_COLS - 4)
    ]
    rows = []
    specials = SPECIALS.get(idx, {})
    for i in range(N_ROWS):
        material, rate = f"Common Material {i % 40}", float(i)
        if i in specials:
            material, rate = specials[i]
        # Filler cells carry realistic weight (~18 chars, like the material
        # names) so the byte budget measures cells, not an artefact of tiny
        # placeholder strings.
        filler = [f"filler value {i:04d}-{j:02d}" for j in range(N_COLS - 4)]
        rows.append([material, "KG", rate, client] + filler)
    return [header] + rows


def _side_grid(client, sheet):
    return [["Client", "Note"], [client, f"{sheet} note"], [client, "n2"]]


class GraphFake:
    """Serves addressed range reads from in-memory grids, with accounting."""

    def __init__(self):
        self.grids = {}
        self.requests = 0
        self.bytes = 0

    def add(self, item_id, sheet, grid):
        self.grids[(item_id, sheet)] = grid

    async def get_range(self, item_id, sheet_name, address, session_id=None, **kw):
        grid = self.grids[(item_id, sheet_name)]
        col1, row1, col2, row2 = ranges.parse_range(address)
        values = []
        for row in grid[row1 - 1 : row2]:
            chunk = list(row[col1 - 1 : col2])
            chunk += [None] * ((col2 - col1 + 1) - len(chunk))
            values.append(chunk)
        self.requests += 1
        self.bytes += len(json.dumps(values))
        formats = [["General"] * len(v) for v in values]
        return {
            "values": values,
            "address": f"{sheet_name}!{address}",
            "numberFormat": formats,
        }

    def full_sheet_bytes(self, item_id, sheet):
        return len(json.dumps(self.grids[(item_id, sheet)]))


def _distinct_strings(grid, col_idx):
    seen = {}
    for row in grid[1:]:
        v = row[col_idx]
        if isinstance(v, str):
            seen[v] = None
    return sorted(seen)


@pytest.fixture
def fake(monkeypatch):
    fake = GraphFake()
    files = {}
    last_col = ranges.index_to_col(N_COLS)
    for idx in range(len(CLIENTS)):
        item_id = f"item{idx}"
        grid = _rate_card_grid(idx)
        fake.add(item_id, "Rate Card", grid)
        sheets = {
            "Rate Card": {
                "header_row": 1,
                "columns": grid[0],
                "used_range_address": f"'Rate Card'!A1:{last_col}{N_ROWS + 1}",
                "row_count": N_ROWS + 1,
                "column_count": N_COLS,
                "sampled_values": {
                    "Material": _distinct_strings(grid, 0),
                    "Unit": ["KG"],
                    "Client": [CLIENTS[idx]],
                },
            }
        }
        for side in ("Orders", "Contacts", "Terms", "Summary"):
            fake.add(item_id, side, _side_grid(CLIENTS[idx], side))
            sheets[side] = {
                "header_row": 1,
                "columns": ["Client", "Note"],
                "used_range_address": f"{side}!A1:B3",
                "row_count": 3,
                "column_count": 2,
                "sampled_values": {"Client": [CLIENTS[idx]]},
            }
        files[_file_name(idx)] = {"item_id": item_id, "sheets": sheets}

    graph = {"workspaces": {"/Contracts": {"schema_version": SCHEMA, "files": files}}}
    monkeypatch.setattr(lookup, "load_graph", lambda: graph)
    monkeypatch.setattr(lookup, "get_range", fake.get_range)
    monkeypatch.setattr(lookup, "search", lambda *a, **k: [])
    return fake


def _lookup(**kwargs):
    kwargs.setdefault("folder_path", "/Contracts")
    return asyncio.run(lookup.execute_lookup(**kwargs))


# ---- acceptance criterion 1: one NL call returns the cell with provenance


def test_natural_language_lookup_returns_cell_with_provenance(fake):
    result = _lookup(
        query="What's the contracted rate for Titanium Dioxide under the BESTEX contract?"
    )
    assert result["found"] is True
    assert result["value"] == 142.5
    assert result["confidence"] == "high"
    prov = result["provenance"]
    assert prov["file"] == "CON000_BESTEX.xlsx"
    assert prov["sheet"] == "Rate Card"
    # Data row 346 (0-based) sits at absolute sheet row 348; Contracted Rate
    # is column C.
    assert prov["cell"] == "C348"
    assert prov["key_column"] == "Material"
    assert prov["return_column"] == "Contracted Rate"
    assert prov["matched_row"]["Material"] == "Titanium Dioxide"
    assert prov["matched_row"]["Unit"] == "KG"


# ---- acceptance criterion 2: request and byte budget


def test_lookup_stays_within_request_and_byte_budget(fake):
    _lookup(
        query="What's the contracted rate for Titanium Dioxide under the BESTEX contract?"
    )
    assert fake.requests <= 8
    full = fake.full_sheet_bytes("item0", "Rate Card")
    assert fake.bytes < 0.05 * full, (
        f"lookup moved {fake.bytes} bytes; whole-sheet fetch is {full}"
    )


# ---- acceptance criterion 3: agreement across workbooks -> high + names all


def test_corroborating_workbooks_agree(fake):
    result = _lookup(query="contracted rate for Calcium Carbonate")
    assert result["found"] is True
    assert result["confidence"] == "high"
    assert result["value"] == 88.0
    found_in = {result["provenance"]["file"]} | {
        c["file"] for c in result["provenance"]["corroborated_by"]
    }
    assert found_in == {_file_name(1), _file_name(2), _file_name(3)}


# ---- acceptance criterion 4: disagreement -> conflict, value null


def test_conflicting_workbooks_force_conflict(fake):
    result = _lookup(query="contracted rate for Zinc Oxide")
    assert result["found"] is True
    assert result["confidence"] == "conflict"
    assert result["value"] is None
    values = {a["value"] for a in result["alternatives"]}
    assert values == {10.0, 20.0, 30.0}
    files = {a["file"] for a in result["alternatives"]}
    assert files == {_file_name(4), _file_name(5), _file_name(6)}


# ---- acceptance criterion 5: two rows in one sheet -> ambiguous, both rows


def test_duplicate_key_within_sheet_is_ambiguous(fake):
    result = _lookup(query="contracted rate for Sodium Chloride")
    assert result["found"] is True
    assert result["confidence"] == "ambiguous"
    assert result["value"] is None
    assert len(result["alternatives"]) == 2
    assert {a["value"] for a in result["alternatives"]} == {5.0, 7.0}


# ---- acceptance criterion 6: misspelled key -> found false + suggestions


def test_misspelled_key_suggests_corrections(fake):
    result = _lookup(query="contracted rate for Titanum Dioxide")
    assert result["found"] is False
    assert result["value"] is None
    assert "Titanium Dioxide" in result["suggestions"]


# ---- explicit triple mode


def test_explicit_triple_lookup(fake):
    result = _lookup(
        key_column="Material",
        key_value="Titanium Dioxide",
        return_column="Contracted Rate",
    )
    assert result["found"] is True
    assert result["value"] == 142.5
    assert result["provenance"]["file"] == "CON000_BESTEX.xlsx"


def test_explicit_lookup_normalises_key_matching(fake):
    result = _lookup(
        key_column="Material",
        key_value="  titanium dioxide ",
        return_column="Contracted Rate",
    )
    assert result["found"] is True
    assert result["value"] == 142.5


def test_explicit_unknown_column_fails_loud(fake):
    with pytest.raises(ValueError, match="not found"):
        _lookup(
            key_column="Materiel",
            key_value="Titanium Dioxide",
            return_column="Contracted Rate",
        )


def test_scope_narrows_candidates(fake):
    result = _lookup(
        query="contracted rate for Zinc Oxide",
        scope={"file": _file_name(5)},
    )
    assert result["found"] is True
    assert result["confidence"] == "high"
    assert result["value"] == 20.0


# ---- routing floor


def test_unroutable_query_reports_routing_ambiguity(fake):
    result = _lookup(query="what is the meaning of life")
    assert result["found"] is False
    assert result["ambiguity"] is not None


# ---- acceptance criterion 8: never a bare value


def test_every_outcome_carries_provenance_or_ambiguity(fake):
    outcomes = [
        _lookup(query="contracted rate for Titanium Dioxide for BESTEX"),
        _lookup(query="contracted rate for Zinc Oxide"),
        _lookup(query="contracted rate for Sodium Chloride"),
        _lookup(query="contracted rate for Titanum Dioxide"),
        _lookup(query="what is the meaning of life"),
    ]
    for result in outcomes:
        has_provenance = bool(result.get("provenance")) or bool(
            result.get("alternatives")
        )
        has_reason = result.get("ambiguity") is not None
        assert has_provenance or has_reason, f"bare response: {result}"


# ---- acceptance criterion 7: get_cell is one request


def test_get_cell_is_one_request(fake):
    result = asyncio.run(
        lookup.execute_get_cell(
            _file_name(0), "Rate Card", "C348", "/Contracts"
        )
    )
    assert fake.requests == 1
    assert result["value"] == 142.5
    assert result["resolved_type"] == "number"
    assert result["address"] == "C348"


def test_get_cell_rejects_multi_cell_addresses(fake):
    with pytest.raises(ValueError, match="single cell"):
        asyncio.run(
            lookup.execute_get_cell(
                _file_name(0), "Rate Card", "A1:C10", "/Contracts"
            )
        )
