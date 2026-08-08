"""Tier 1 region detection, measured against the reference workbook's geometry.

The numbers in ``REFERENCE_SHEETS`` are the real row spans of
SMC-SLS-BESTEX-26-27.xlsx — eight sheets, eleven table bodies, four sections on
one sheet and two identically-named tables on another. Only the geometry is
kept; the figures and the client names are not this workbook's.

The acceptance bar for Tier 1: all eleven bodies found, and every header either
exact or one row late — never further. The four that miss each open with an
``Opening Balance`` line between the header and the first numbered row, which
no formula can reveal and only Tier 2 can settle.
"""

import asyncio

import pytest

from excelmcp import regions, structure
from excelmcp.regions import (
    build_regions,
    derive_sheet_regions,
    extract_body_spans,
    extract_sheet_references,
    merge_spans,
    unclaimed_rows,
)

# sheet -> (last used row, [(body_start, body_end)], [real header row per body])
REFERENCE_SHEETS: dict[str, tuple[int, list[tuple[int, int]], list[int]]] = {
    "Dashboard": (40, [], []),
    "Batch Input & Output": (620, [(16, 616)], [15]),
    "Despatch Register": (60, [(16, 55)], [15]),
    "CMC Statement": (50, [(7, 46)], [5]),
    "Methanol Statement": (50, [(7, 46)], [5]),
    "H2SO4 & Soda Ash": (82, [(7, 31), (48, 72)], [5, 46]),
    "Job Work Costing": (60, [(15, 18)], [14]),
    "Other Issues": (
        62,
        [(6, 15), (20, 29), (34, 43), (48, 57)],
        [5, 19, 33, 47],
    ),
}


def totals_formulas(bodies: list[tuple[int, int]]) -> list[list[str]]:
    """A formulas grid shaped like a real sheet's: a totals row under each body.

    Each body gets the mix these sheets actually carry — a couple of column
    sums, a COUNTA, and on the second column a sum split either side of a
    one-row gap, which only counts as one body once the spans are merged.
    """
    grid: list[list[str]] = []
    for start, end in bodies:
        mid = start + (end - start) // 2
        grid.append(
            [
                f"=SUM(C{start}:C{end})",
                f"=SUM(D{start}:D{mid})+SUM(D{mid + 2}:D{end})",
                f"=COUNTA(B{start}:B{end})",
            ]
        )
    return grid


def test_every_reference_body_is_found():
    """Test 1 of the brief: 11 of 11 bodies across the 8 sheets."""
    total = 0
    for sheet, (last_row, bodies, _) in REFERENCE_SHEETS.items():
        found = merge_spans(extract_body_spans(totals_formulas(bodies)))
        assert found == bodies, f"{sheet}: got {found}, want {bodies}"
        total += len(found)
    assert total == 11


def test_header_guess_is_exact_on_seven_and_within_one_on_four():
    """Test 2 of the brief: the residual ambiguity is one row, and no wider.

    The brief's prose says eight exact and three off by one; its own measurement
    table says otherwise, because 'H2SO4 & Soda Ash' has *two* bodies and both
    of them miss by a row. Seven and four is what the table adds up to, and it
    is what the geometry produces. The claim that matters is unchanged: every
    miss is exactly one row, caused by an Opening Balance line between the
    header and the first numbered row.
    """
    exact = 0
    off_by_one = 0
    for sheet, (last_row, bodies, real_headers) in REFERENCE_SHEETS.items():
        found = build_regions(
            merge_spans(extract_body_spans(totals_formulas(bodies))), 1, last_row
        )
        assert len(found) == len(real_headers)
        for region, real in zip(found, real_headers):
            delta = region["header_row"] - real
            assert 0 <= delta <= 1, f"{sheet}: header off by {delta}"
            exact += delta == 0
            off_by_one += delta == 1
    assert (exact, off_by_one) == (7, 4)
    assert exact + off_by_one == 11


def test_unclaimed_rows_cover_every_summary_and_title_block():
    """Test 3: the gaps the server admits it cannot classify.

    On the two-table sheet those are the banner above the first section, the
    block between the sections, and the closing summary — precisely where the
    paired-column key/value data lives.
    """
    last_row, bodies, _ = REFERENCE_SHEETS["H2SO4 & Soda Ash"]
    found, gaps = derive_sheet_regions(totals_formulas(bodies), 1, last_row)
    assert [r["body"] for r in found] == ["7:31", "48:72"]
    assert gaps == ["1:5", "32:46", "73:82"]


def test_multi_section_sheet_keeps_its_sections_apart():
    """Four sections with identical columns must not merge into one body."""
    last_row, bodies, _ = REFERENCE_SHEETS["Other Issues"]
    found, _ = derive_sheet_regions(totals_formulas(bodies), 1, last_row)
    assert [r["body"] for r in found] == ["6:15", "20:29", "34:43", "48:57"]


class TestSpanExtraction:
    def test_only_aggregate_functions_contribute(self):
        # A VLOOKUP's table array is not a body the sheet totals.
        assert extract_body_spans([["=VLOOKUP(A1,C5:F90,2,FALSE)"]]) == []

    def test_criteria_ranges_of_sumifs_count(self):
        spans = extract_body_spans([['=SUMIFS(D7:D31,B7:B31,"Receipt")']])
        assert spans == [(7, 31)]

    def test_commas_inside_string_literals_do_not_split_arguments(self):
        spans = extract_body_spans([['=SUMIF(B7:B31,">1,000",D7:D31)']])
        assert spans == [(7, 31)]

    def test_nested_calls_are_parsed(self):
        spans = extract_body_spans([["=ROUND(SUM(C10:C40)/COUNTA(B10:B40),2)"]])
        assert spans == [(10, 40)]

    def test_anchored_references_are_the_same_span(self):
        assert extract_body_spans([["=SUM($C$7:$C$31)"]]) == [(7, 31)]

    def test_cross_sheet_ranges_are_not_local_bodies(self):
        assert extract_body_spans([["=SUM('CMC Statement'!C7:C46)"]]) == []

    def test_full_column_references_are_ignored(self):
        assert extract_body_spans([["=SUM(C:C)"]]) == []

    def test_non_formula_cells_are_skipped(self):
        assert extract_body_spans([["SUM(C7:C31)", 42, None, ""]]) == []

    def test_function_name_must_stand_alone(self):
        # MYSUM(...) is not SUM(...).
        assert extract_body_spans([["=MYSUM(C7:C31)"]]) == []


class TestSpanMerging:
    def test_short_spans_are_not_tables(self):
        assert merge_spans([(3, 5)]) == []

    def test_a_gap_of_two_still_merges(self):
        assert merge_spans([(7, 20), (23, 31)]) == [(7, 31)]

    def test_a_gap_of_three_is_a_separate_section(self):
        assert merge_spans([(7, 20), (24, 31)]) == [(7, 20), (24, 31)]

    def test_overlapping_spans_collapse(self):
        assert merge_spans([(7, 31), (7, 25), (10, 31)]) == [(7, 31)]

    def test_a_nested_span_does_not_shrink_its_parent(self):
        assert merge_spans([(7, 40), (10, 15)]) == [(7, 40)]


class TestRegionBounds:
    def test_a_body_at_the_top_has_no_header_row(self):
        region = build_regions([(1, 20)], 1, 30)[0]
        assert region["header_row"] is None
        assert region["range"] == "1:20"

    def test_a_body_past_the_used_range_is_clamped(self):
        assert build_regions([(7, 900)], 1, 50)[0]["body"] == "7:50"

    def test_a_body_entirely_outside_the_used_range_is_dropped(self):
        assert build_regions([(80, 120)], 1, 50) == []

    def test_regions_are_flagged_unconfirmed(self):
        assert build_regions([(7, 31)], 1, 40)[0]["confidence"] == "unconfirmed"

    def test_a_sheet_with_no_formulas_is_entirely_unclaimed(self):
        found, gaps = derive_sheet_regions([], 1, 40)
        assert (found, gaps) == ([], ["1:40"])

    def test_unclaimed_rows_of_a_fully_covered_sheet_are_empty(self):
        assert unclaimed_rows(build_regions([(2, 40)], 1, 40), 1, 40) == []


class TestSheetReferences:
    """Test 10: cross-sheet formula references, recorded as stated facts."""

    def test_quoted_and_bare_sheet_names_both_resolve(self):
        found = extract_sheet_references(
            [["='CMC Statement'!C48", "=Dashboard!B2+1"]]
        )
        assert found == [
            {"sheet": "CMC Statement", "addresses": ["C48"]},
            {"sheet": "Dashboard", "addresses": ["B2"]},
        ]

    def test_addresses_accumulate_per_sheet_and_deduplicate(self):
        found = extract_sheet_references(
            [["='CMC Statement'!C48+'CMC Statement'!C49"], ["='CMC Statement'!C48"]]
        )
        assert found == [
            {"sheet": "CMC Statement", "addresses": ["C48", "C49"]}
        ]

    def test_external_workbook_references_are_dropped(self):
        assert extract_sheet_references([["=[1]Sheet1!A1"]]) == []

    def test_anchors_are_stripped_from_the_address(self):
        found = extract_sheet_references([["='Job Work Costing'!$C$15:$C$18"]])
        assert found[0]["addresses"] == ["C15:C18"]

    def test_apostrophes_in_a_sheet_name_survive(self):
        found = extract_sheet_references([["='Bob''s Sheet'!A1"]])
        assert found[0]["sheet"] == "Bob's Sheet"


class TestFormulaRelationships:
    """A reference edge is a fact, but it is not a join key."""

    def _files(self):
        return {
            "book.xlsx": {
                "sheets": {
                    "Dashboard": {
                        "used_range_address": "Dashboard!A1:D40",
                        "columns": ["Metric", "Value"],
                        "sheet_references": [
                            {"sheet": "CMC Statement", "addresses": ["C48"]}
                        ],
                    },
                    "CMC Statement": {
                        "used_range_address": "CMC Statement!A1:E50",
                        "columns": ["Date", "Party", "Qty (Kgs)", "Rate", "Value"],
                    },
                }
            }
        }

    def test_reference_carries_full_confidence_and_its_own_kind(self):
        rel = structure.formula_relationships(self._files())[0]
        assert rel["source"] == "formula"
        assert rel["kind"] == "reference"
        assert rel["confidence"] == 1.0
        assert rel["left"]["sheet"] == "Dashboard"
        assert rel["right"]["sheet"] == "CMC Statement"

    def test_the_target_column_is_resolved_from_the_address(self):
        rel = structure.formula_relationships(self._files())[0]
        assert rel["right"]["column"] == "Qty (Kgs)"

    def test_a_reference_never_becomes_a_join_key(self):
        from excelmcp.query_engine import _suggest_join_keys

        workspace = {"files": self._files(), "relationships": []}
        workspace["relationships"] = structure.formula_relationships(self._files())
        with pytest.raises(ValueError, match="No confident relationship"):
            _suggest_join_keys(
                "/ERP",
                workspace,
                "book.xlsx",
                "Dashboard",
                "book.xlsx",
                "CMC Statement",
            )

    def test_a_reference_to_an_unknown_sheet_is_dropped(self):
        files = self._files()
        del files["book.xlsx"]["sheets"]["CMC Statement"]
        assert structure.formula_relationships(files) == []

    def test_a_self_reference_is_not_a_relationship(self):
        files = self._files()
        files["book.xlsx"]["sheets"]["Dashboard"]["sheet_references"] = [
            {"sheet": "Dashboard", "addresses": ["A1"]}
        ]
        assert structure.formula_relationships(files) == []


class TestColumnNameCleanup:
    def test_wrapped_header_cells_collapse_to_one_line(self):
        assert structure.columns_from_row(["Qty\n(Kgs)", "Opening\nBalance"]) == [
            "Qty (Kgs)",
            "Opening Balance",
        ]

    def test_the_exact_spelling_stays_available(self):
        assert structure.raw_columns_from_row(["Qty\n(Kgs)"]) == ["Qty\n(Kgs)"]

    def test_raw_and_cleaned_names_stay_positionally_aligned(self):
        row = ["Qty\n(Kgs)", "", "Qty\n(Kgs)"]
        cleaned = structure.columns_from_row(row)
        raw = structure.raw_columns_from_row(row)
        assert len(cleaned) == len(raw) == 3
        assert cleaned == ["Qty (Kgs)", "Column_1", "Qty (Kgs)_1"]
        assert raw == ["Qty\n(Kgs)", "Column_1", "Qty\n(Kgs)_1"]


class TestUnicodeNormalisation:
    """Test 9: a query saying H2SO4 must reach 'Sulphuric Acid (H₂SO₄)'."""

    def test_subscripts_fold_to_digits(self):
        assert structure.normalise_name("Sulphuric Acid (H₂SO₄)") == (
            "sulphuric acid (h2so4)"
        )

    def test_superscripts_fold_to_digits(self):
        assert structure.normalise_name("m³") == "m3"

    def test_both_spellings_of_a_sheet_name_agree(self):
        assert structure.normalise_name("H₂SO₄ & Soda Ash") == structure.normalise_name(
            "h2so4 & soda ash"
        )

    def test_fuzzy_candidates_reach_across_the_two_spellings(self):
        assert structure.fuzzy_name_candidates(
            "H2SO4 & Soda Ash", ["H₂SO₄ & Soda Ash", "Methanol Statement"]
        ) == ["H₂SO₄ & Soda Ash"]


class TestSchemaGuard:
    def test_a_pre_v3_workspace_is_refused_by_name(self):
        with pytest.raises(ValueError, match="predates v0.3"):
            structure.require_current_schema({"files": {}}, "/ERP")

    def test_the_message_names_the_command_that_fixes_it(self):
        with pytest.raises(ValueError, match="Run scan_workspace on '/ERP'"):
            structure.require_current_schema({"schema_version": 2}, "/ERP")

    def test_a_current_workspace_passes(self):
        structure.require_current_schema(
            {"schema_version": structure.GRAPH_SCHEMA_VERSION}, "/ERP"
        )


class TestScanUsesFormulaHeaders:
    """Test 11 in reverse: the formula header must win where it disagrees, and
    the clean single-region path must behave exactly as it did in v0.2."""

    def _run(self, monkeypatch, meta, ranges, formulas):
        async def fake_used_range(item_id, sheet, session=None, *, select=None, **kw):
            if select and "formulas" in select:
                return {"formulas": formulas}
            if select and "values" not in select:
                return meta
            return {"values": []}

        async def fake_get_range(item_id, sheet, address, session=None, **kw):
            return ranges.get(address, {"values": []})

        monkeypatch.setattr(structure, "get_used_range", fake_used_range)
        monkeypatch.setattr(structure, "get_range", fake_get_range)
        return asyncio.run(structure._scan_sheet_structure("i", "S", None))

    def test_a_banner_no_longer_becomes_the_header(self, monkeypatch):
        # Rows 1-4 are a client banner the heuristic would seize on; the real
        # header is row 5 and the formulas total rows 6..15.
        entry = self._run(
            monkeypatch,
            meta={"address": "S!A1:C20", "rowCount": 20, "columnCount": 3},
            ranges={
                "A1:C10": {
                    "values": [
                        ["SLS-BESTEX", "Contract", ""],
                        ["", "", ""],
                        ["", "", ""],
                        ["", "", ""],
                        ["Date", "Party", "Qty\n(Kgs)"],
                        ["2026-01-01", "Acme", 10],
                    ]
                },
                "A6:C15": {"values": [["2026-01-01", "Acme", 10]]},
            },
            formulas=[["=SUM(C6:C15)"]],
        )
        assert entry["header_row"] == 5
        assert entry["header_source"] == "formula"
        assert entry["columns"] == ["Date", "Party", "Qty (Kgs)"]
        assert entry["columns_raw"] == ["Date", "Party", "Qty\n(Kgs)"]
        assert entry["regions"][0]["body"] == "6:15"
        assert entry["unclaimed_rows"] == ["1:4", "16:20"]
        assert entry["layout_confidence"] == "unconfirmed"

    def test_a_clean_single_region_sheet_is_unchanged(self, monkeypatch):
        entry = self._run(
            monkeypatch,
            meta={"address": "S!A1:C500", "rowCount": 500, "columnCount": 3},
            ranges={
                "A1:C10": {"values": [["Name", "Qty", "Status"], ["a", 1, "Open"]]},
                "A2:C500": {"values": [["a", 1, "Open"], ["b", 2, "Closed"]]},
            },
            formulas=[["=SUM(B2:B500)"]],
        )
        assert entry["header_row"] == 1
        assert entry["columns"] == ["Name", "Qty", "Status"]
        assert entry["approx_row_count"] == 499
        assert entry["sampled_values"]["Status"] == ["Closed", "Open"]
        assert "columns_raw" not in entry

    def test_no_formulas_falls_back_to_the_heuristic(self, monkeypatch):
        entry = self._run(
            monkeypatch,
            meta={"address": "S!A1:B20", "rowCount": 20, "columnCount": 2},
            ranges={
                "A1:B10": {"values": [["Name", "Qty"], ["a", 1]]},
                "A2:B20": {"values": [["a", 1]]},
            },
            formulas=[],
        )
        assert entry["header_source"] == "heuristic"
        assert entry["header_row"] == 1
        assert entry["regions"] == []
        assert entry["unclaimed_rows"] == ["1:20"]

    def test_a_failed_formula_read_does_not_abort_the_scan(self, monkeypatch):
        async def failing(item_id, sheet, session=None, *, select=None, **kw):
            if select and "formulas" in select:
                raise RuntimeError("Graph said no")
            if select and "values" not in select:
                return {"address": "S!A1:B20", "rowCount": 20, "columnCount": 2}
            return {"values": []}

        async def fake_get_range(item_id, sheet, address, session=None, **kw):
            return {"values": [["Name", "Qty"], ["a", 1]]}

        monkeypatch.setattr(structure, "get_used_range", failing)
        monkeypatch.setattr(structure, "get_range", fake_get_range)
        entry = asyncio.run(structure._scan_sheet_structure("i", "S", None))
        assert entry["columns"] == ["Name", "Qty"]
        assert entry["regions"] == []

    def test_sampling_stops_at_the_end_of_the_first_body(self, monkeypatch):
        """Section titles below a body must never become sampled 'values'."""
        addresses = []

        async def fake_used_range(item_id, sheet, session=None, *, select=None, **kw):
            if select and "formulas" in select:
                return {"formulas": [["=SUM(C6:C15)"], ["=SUM(C20:C29)"]]}
            if select and "values" not in select:
                return {"address": "S!A1:C60", "rowCount": 60, "columnCount": 3}
            return {"values": []}

        async def fake_get_range(item_id, sheet, address, session=None, **kw):
            addresses.append(address)
            if address == "A1:C10":
                return {
                    "values": [["", "", ""]] * 4
                    + [["Date", "Party", "Qty"]]
                    + [["2026-01-01", "Acme", 1]] * 5
                }
            return {"values": [["2026-01-01", "Acme", 1]]}

        monkeypatch.setattr(structure, "get_used_range", fake_used_range)
        monkeypatch.setattr(structure, "get_range", fake_get_range)
        entry = asyncio.run(structure._scan_sheet_structure("i", "S", None))
        assert len(entry["regions"]) == 2
        # Not A6:C60 — the sample stops where the first body does.
        assert "A6:C15" in addresses
