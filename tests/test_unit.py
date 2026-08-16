"""Offline unit tests — no network, no OneDrive, no real ~/.excelmcp writes.

Every test that touches the config directory redirects the module-level path
constants at a tmp_path first, so running the suite can never clobber a real
user's token cache or search index.
"""

import asyncio
import json

import msal
import numpy as np
import pandas as pd
import pytest

from excelmcp import auth, embeddings, graph_client, query_engine, ranges, storage, structure
from excelmcp.structure import GRAPH_SCHEMA_VERSION as SCHEMA
from excelmcp.graph_client import (
    _parse_retry_after,
    quote_drive_path,
    quote_odata_literal,
)


# ---------------------------------------------------------------- URL encoding


class TestUrlEncoding:
    def test_apostrophe_is_odata_escaped_then_percent_encoded(self):
        # A single quote must be doubled for OData before percent-encoding,
        # otherwise it terminates the literal in worksheets('...').
        assert quote_odata_literal("Bob's Data") == "Bob%27%27s%20Data"

    def test_slash_cannot_inject_a_path_segment(self):
        # The old code used quote() with its default safe="/", so this escaped
        # the worksheets('...') literal and addressed a different resource.
        assert "/" not in quote_odata_literal("Q1/Q2")
        assert quote_odata_literal("Q1/Q2") == "Q1%2FQ2"

    def test_plain_sheet_name_is_untouched_except_spaces(self):
        assert quote_odata_literal("Sheet1") == "Sheet1"

    @pytest.mark.parametrize("hostile", ["a'b", "a/b", "a?b", "a#b", "a%b", "'"])
    def test_no_url_metacharacters_survive(self, hostile):
        encoded = quote_odata_literal(hostile)
        assert not set(encoded) & set("/?#'%") - {"%"}
        # '%' only ever appears as the start of an escape triple.
        for i, ch in enumerate(encoded):
            if ch == "%":
                assert len(encoded) >= i + 3

    def test_drive_path_encodes_colon(self):
        # ':' delimits the /root:<path>: addressing form.
        assert ":" not in quote_drive_path("/a:b/c")

    def test_drive_path_keeps_separators(self):
        assert quote_drive_path("/ERP/2026") == "/ERP/2026"

    def test_drive_path_normalises_empty_segments(self):
        assert quote_drive_path("//ERP//") == "/ERP"


# ---------------------------------------------------------- range requests


@pytest.fixture
def captured_requests(monkeypatch):
    """Replaces the HTTP layer, recording every (method, url, params) triple."""
    calls = []

    async def fake_request(method, url, headers=None, json_data=None, params=None):
        calls.append({"method": method, "url": url, "params": params})
        return {"values": []}

    monkeypatch.setattr(graph_client, "_request", fake_request)
    return calls


class TestRangeRequests:
    def test_used_range_selects_only_data_fields(self, captured_requests):
        # Without $select Graph returns text/formulas/numberFormat/... — several
        # full-size 2D arrays of which only `values` was ever read.
        asyncio.run(graph_client.get_used_range("item", "Sheet1"))
        call = captured_requests[0]
        assert call["url"].endswith("/usedRange")
        assert call["params"] == {"$select": "values,address,rowCount,columnCount"}

    def test_values_only_variant(self, captured_requests):
        asyncio.run(graph_client.get_used_range("item", "Sheet1", values_only=True))
        assert captured_requests[0]["url"].endswith("/usedRange(valuesOnly=true)")

    def test_select_none_fetches_full_resource(self, captured_requests):
        asyncio.run(graph_client.get_used_range("item", "Sheet1", select=None))
        assert captured_requests[0]["params"] is None

    def test_get_range_addresses_one_range(self, captured_requests):
        asyncio.run(graph_client.get_range("item", "Sheet1", "B7"))
        call = captured_requests[0]
        assert "range(address='B7')" in call["url"]
        assert call["params"] == {"$select": "values,address,rowCount,columnCount"}

    def test_get_range_encodes_hostile_addresses(self, captured_requests):
        # An address is caller-supplied text; it must not break out of the URL.
        asyncio.run(graph_client.get_range("item", "Q1/Q2", "A1"))
        assert "Q1/Q2" not in captured_requests[0]["url"]


# ------------------------------------------------------------- Retry-After


class TestRetryAfter:
    def test_integer_seconds(self):
        assert _parse_retry_after("30", 0) == 30.0

    def test_http_date_does_not_crash(self):
        # int("Wed, 21 Oct 2015 07:28:00 GMT") used to raise ValueError and kill
        # the request instead of retrying.
        result = _parse_retry_after("Wed, 21 Oct 2015 07:28:00 GMT", 0)
        assert isinstance(result, float)
        assert 1.0 <= result <= 60.0

    def test_garbage_falls_back_to_backoff(self):
        assert _parse_retry_after("not-a-header", 0) == 2.0

    def test_missing_header_falls_back(self):
        assert _parse_retry_after(None, 1) == 4.0

    def test_absurd_value_is_clamped(self):
        assert _parse_retry_after("999999", 0) == 60.0


# --------------------------------------------------------------- conditions


def _frame():
    return pd.DataFrame(
        {
            "Name": ["alpha", "a.pha", "beta", "gamma"],
            "Qty": ["10", "20", "300", "abc"],
        }
    )


class TestConditions:
    def test_contains_is_literal_not_regex(self):
        # "a.pha" must match only itself; as a regex it would also match "alpha".
        out = query_engine._apply_conditions(_frame(), {"Name": "~a.pha"})
        assert list(out["Name"]) == ["a.pha"]

    def test_regex_metacharacters_do_not_explode(self):
        # A catastrophic-backtracking pattern is just a literal substring now.
        out = query_engine._apply_conditions(_frame(), {"Name": "~(a+)+$"})
        assert len(out) == 0

    def test_greater_than(self):
        out = query_engine._apply_conditions(_frame(), {"Qty": ">15"})
        assert sorted(out["Qty"]) == ["20", "300"]

    def test_greater_or_equal_is_supported(self):
        # ">=20" previously parsed as ">" with bound "=20", failed float(), and
        # was silently dropped — returning every row as if unfiltered.
        out = query_engine._apply_conditions(_frame(), {"Qty": ">=20"})
        assert sorted(out["Qty"]) == ["20", "300"]

    def test_less_or_equal_is_supported(self):
        out = query_engine._apply_conditions(_frame(), {"Qty": "<=20"})
        assert sorted(out["Qty"]) == ["10", "20"]

    def test_bad_numeric_bound_raises_instead_of_silently_passing(self):
        with pytest.raises(ValueError, match="neither a number nor an ISO date"):
            query_engine._apply_conditions(_frame(), {"Qty": ">abc"})

    def test_unknown_column_raises(self):
        with pytest.raises(ValueError, match="not found"):
            query_engine._apply_conditions(_frame(), {"Nope": "x"})

    def test_result_is_a_copy_not_a_view(self):
        out = query_engine._apply_conditions(_frame(), {"Qty": ">15"})
        out["Qty"] = pd.to_numeric(out["Qty"])  # must not warn or fail
        assert out["Qty"].tolist() == [20, 300]


def _status_frame():
    return pd.DataFrame(
        {
            "Status": ["Closed", "Closed ", "closed", "Open"],
            "Qty": [1, 2, 3, 4],
        }
    )


class TestNormalisedMatching:
    def test_exact_match_ignores_case_and_whitespace(self):
        # {"Status": "closed"} returning zero rows against a sheet containing
        # "Closed " made agents report "no closed orders" with full confidence.
        out = query_engine._apply_conditions(_status_frame(), {"Status": "closed"})
        assert len(out) == 3

    def test_exact_case_flag_restores_strict_matching(self):
        out = query_engine._apply_conditions(
            _status_frame(), {"Status": "Closed"}, exact_case=True
        )
        assert len(out) == 1
        assert out["Qty"].tolist() == [1]

    def test_numeric_equality_still_works(self):
        out = query_engine._apply_conditions(_status_frame(), {"Qty": 2})
        assert out["Status"].tolist() == ["Closed "]


class TestZeroMatchDiagnostics:
    def test_offending_column_lists_actual_values(self):
        diag = query_engine._zero_match_diagnostics(
            _status_frame(), {"Status": "Shipped"}
        )
        (entry,) = diag["conditions"]
        assert entry["column"] == "Status"
        assert entry["rows_matching_this_condition_alone"] == 0
        assert "Closed" in entry["distinct_values_present"]
        assert "Open" in entry["distinct_values_present"]

    def test_individually_matching_conditions_show_their_counts(self):
        # Each condition matches alone; only the combination is empty.
        diag = query_engine._zero_match_diagnostics(
            _status_frame(), {"Status": "Open", "Qty": "<4"}
        )
        by_col = {e["column"]: e for e in diag["conditions"]}
        assert by_col["Status"]["rows_matching_this_condition_alone"] == 1
        assert by_col["Qty"]["rows_matching_this_condition_alone"] == 3
        assert "distinct_values_present" not in by_col["Status"]

    def test_distinct_values_are_capped_at_twenty(self):
        df = pd.DataFrame({"Code": [f"C{i}" for i in range(50)]})
        diag = query_engine._zero_match_diagnostics(df, {"Code": "nope"})
        assert len(diag["conditions"][0]["distinct_values_present"]) == 20


# ------------------------------------------------------ concurrency bound


class TestMaxConcurrency:
    def test_default_is_eight(self, monkeypatch):
        monkeypatch.delenv("EXCELMCP_MAX_CONCURRENCY", raising=False)
        assert graph_client.max_concurrency() == 8

    def test_env_override(self, monkeypatch):
        monkeypatch.setenv("EXCELMCP_MAX_CONCURRENCY", "16")
        assert graph_client.max_concurrency() == 16

    def test_garbage_falls_back(self, monkeypatch):
        monkeypatch.setenv("EXCELMCP_MAX_CONCURRENCY", "lots")
        assert graph_client.max_concurrency() == 8

    def test_nonpositive_is_clamped_to_one(self, monkeypatch):
        monkeypatch.setenv("EXCELMCP_MAX_CONCURRENCY", "0")
        assert graph_client.max_concurrency() == 1

    def test_requests_in_flight_never_exceed_the_bound(self, monkeypatch):
        # 100 files used to mean 100 simultaneous usedRange requests, which
        # Graph throttles hard enough to exhaust the retry ladder.
        monkeypatch.setenv("EXCELMCP_MAX_CONCURRENCY", "2")
        monkeypatch.setattr(graph_client, "_request_semaphore", None)
        monkeypatch.setattr(graph_client, "_request_semaphore_loop", None)

        async def fake_token(*args, **kwargs):
            return "tok"

        async def fake_session():
            return object()

        state = {"in_flight": 0, "peak": 0}

        async def fake_locked(session, method, url, headers, json_data, params, token):
            state["in_flight"] += 1
            state["peak"] = max(state["peak"], state["in_flight"])
            await asyncio.sleep(0.005)
            state["in_flight"] -= 1
            return {}

        monkeypatch.setattr(graph_client, "get_token", fake_token)
        monkeypatch.setattr(graph_client, "_get_session", fake_session)
        monkeypatch.setattr(graph_client, "_request_locked", fake_locked)

        async def run():
            await asyncio.gather(
                *[graph_client._request("GET", "http://x") for _ in range(12)]
            )

        asyncio.run(run())
        assert state["peak"] <= 2
        # Reset so later tests build a fresh semaphore on their own loop.
        graph_client._request_semaphore = None
        graph_client._request_semaphore_loop = None


# ------------------------------------------------------ cross-file folding


def _fold_workspace(n_files):
    files = {
        f"W{i}.xlsx": {
            "item_id": f"w{i}",
            "sheets": {"Data": {"header_row": 1, "columns": ["Type", "Qty"]}},
        }
        for i in range(n_files)
    }
    return {"workspaces": {"/ERP": {"schema_version": SCHEMA, "files": files}}}


class TestCrossFileFold:
    """The incremental fold must agree with what pd.concat used to produce."""

    def _patch(self, monkeypatch, frames_by_item, n_files):
        monkeypatch.setattr(
            query_engine, "load_graph", lambda: _fold_workspace(n_files)
        )

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            frame = frames_by_item[item_id]
            if isinstance(frame, Exception):
                raise frame
            return frame.copy(), None

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _run(self, operation, value_col="Qty"):
        return asyncio.run(
            query_engine.execute_cross_file_aggregate(
                "/ERP", "Data", value_col, operation
            )
        )

    def test_sum_folds_across_files(self, monkeypatch):
        self._patch(
            monkeypatch,
            {
                "w0": pd.DataFrame({"Type": ["a"], "Qty": [10]}),
                "w1": pd.DataFrame({"Type": ["b", "b"], "Qty": [20, 30]}),
            },
            2,
        )
        result = self._run("sum")
        assert result["total"] == 60.0
        assert result["per_file"] == {"W0.xlsx": 10.0, "W1.xlsx": 50.0}
        assert result["row_count"] == 3

    def test_mean_stays_row_weighted(self, monkeypatch):
        # (10 + 0+0+0) / 4 = 2.5 — NOT the mean of per-file means (5.0).
        self._patch(
            monkeypatch,
            {
                "w0": pd.DataFrame({"Type": ["a"], "Qty": [10]}),
                "w1": pd.DataFrame({"Type": ["b"] * 3, "Qty": [0, 0, 0]}),
            },
            2,
        )
        assert self._run("mean")["total"] == 2.5

    def test_min_max_ignore_non_numeric(self, monkeypatch):
        self._patch(
            monkeypatch,
            {
                "w0": pd.DataFrame({"Type": ["a"], "Qty": ["oops"]}),
                "w1": pd.DataFrame({"Type": ["b", "b"], "Qty": [7, 2]}),
            },
            2,
        )
        assert self._run("min")["total"] == 2.0
        assert self._run("max")["total"] == 7.0

    def test_count_counts_non_null_only(self, monkeypatch):
        self._patch(
            monkeypatch,
            {
                "w0": pd.DataFrame({"Type": ["a", "a"], "Qty": [1, None]}),
                "w1": pd.DataFrame({"Type": ["b"], "Qty": [3]}),
            },
            2,
        )
        assert self._run("count")["total"] == 2.0

    def test_one_failed_file_is_skipped_and_flagged(self, monkeypatch):
        self._patch(
            monkeypatch,
            {
                "w0": pd.DataFrame({"Type": ["a"], "Qty": [10]}),
                "w1": RuntimeError("network down"),
            },
            2,
        )
        result = self._run("sum")
        assert result["total"] == 10.0
        assert [f["file"] for f in result["skipped_files"]] == ["W1.xlsx"]
        assert "skipped" in result["warning"]

    def test_all_files_failing_raises(self, monkeypatch):
        self._patch(
            monkeypatch,
            {"w0": RuntimeError("x"), "w1": RuntimeError("y")},
            2,
        )
        with pytest.raises(Exception, match="All files failed"):
            self._run("sum")

    def test_column_missing_everywhere_raises(self, monkeypatch):
        self._patch(
            monkeypatch,
            {"w0": pd.DataFrame({"Other": [1]}), "w1": pd.DataFrame({"Other": [2]})},
            2,
        )
        with pytest.raises(ValueError, match="not found in any file"):
            self._run("sum")


# ------------------------------------------------- cross-file completeness


def _variant_workspace():
    """Five files: three name the sheet 'Sales', two name it 'Sales 2024'."""
    files = {}
    for i in range(3):
        files[f"F{i}.xlsx"] = {
            "item_id": f"id{i}",
            "sheets": {"Sales": {"header_row": 1, "columns": ["Client", "Amount"]}},
        }
    for i in range(3, 5):
        files[f"F{i}.xlsx"] = {
            "item_id": f"id{i}",
            "sheets": {"Sales 2024": {"header_row": 1, "columns": ["Client", "Amount"]}},
        }
    return {"workspaces": {"/ERP": {"schema_version": SCHEMA, "files": files}}}


class TestCrossFileCompleteness:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        monkeypatch.setattr(query_engine, "load_graph", _variant_workspace)

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            return pd.DataFrame({"Client": ["A", "B"], "Amount": [100, 100]}), None

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _run(self, sheet):
        return asyncio.run(
            query_engine.execute_cross_file_aggregate("/ERP", sheet, "Amount", "sum")
        )

    def test_unmatched_files_are_reported_not_hidden(self):
        # These two files were previously excluded with no signal at all —
        # the README's "visibly partial" claim was false.
        result = self._run("Sales")
        assert [u["file"] for u in result["unmatched_files"]] == ["F3.xlsx", "F4.xlsx"]
        assert all(u["sheets"] == ["Sales 2024"] for u in result["unmatched_files"])

    def test_warning_names_the_excluded_files(self):
        result = self._run("Sales")
        assert result["warning"] is not None
        assert "NOT included" in result["warning"]
        assert "F3.xlsx" in result["warning"] and "F4.xlsx" in result["warning"]

    def test_fuzzy_matches_are_suggested_never_included(self):
        result = self._run("Sales")
        assert result["total"] == 600.0  # never 1000 — suggestions don't count
        for u in result["unmatched_files"]:
            assert "Sales 2024" in u["did_you_mean"]

    def test_exact_match_everywhere_yields_no_warning(self):
        result = self._run("Sales 2024")
        # Files with only 'Sales' are unmatched, but the ones asked about work.
        assert result["total"] == 400.0
        assert len(result["unmatched_files"]) == 3

    def test_no_match_anywhere_raises_with_candidates(self):
        with pytest.raises(ValueError) as exc:
            self._run("Salez")
        assert "Sales" in str(exc.value)


class TestNameNormalisation:
    def test_case_and_whitespace_collapse(self):
        assert structure.normalise_name("  Sales  2024 ") == "sales 2024"
        assert structure.normalise_name("SALES") == structure.normalise_name("sales ")

    def test_fuzzy_catches_containment(self):
        assert "Sales 2024" in structure.fuzzy_name_candidates(
            "Sales", ["Sales 2024", "Costs"]
        )

    def test_fuzzy_catches_small_typos(self):
        assert "Sales" in structure.fuzzy_name_candidates("Slaes", ["Sales", "Costs"])

    def test_fuzzy_rejects_unrelated_names(self):
        assert structure.fuzzy_name_candidates("Sales", ["Inventory", "HR"]) == []

    def test_variants_groups_only_fragmented_names(self):
        files = {
            "a.xlsx": {"sheets": {"Sales": {}}},
            "b.xlsx": {"sheets": {"sales ": {}}},
            "c.xlsx": {"sheets": {"Costs": {}}},
        }
        variants = structure.sheet_name_variants(files)
        assert set(variants) == {"sales"}
        assert variants["sales"] == {"Sales": ["a.xlsx"], "sales ": ["b.xlsx"]}


# ------------------------------------------------------------ JSON safety


class TestRecords:
    def test_nan_becomes_none(self):
        df = pd.DataFrame({"a": [1.0, np.nan]})
        records = query_engine._records(df)
        assert records[1]["a"] is None
        json.dumps(records)  # would raise on a bare NaN

    def test_numpy_scalars_become_python_natives(self):
        df = pd.DataFrame({"a": np.array([1, 2], dtype=np.int64)})
        records = query_engine._records(df)
        assert all(type(r["a"]) is int for r in records)
        json.dumps(records)

    def test_infinity_becomes_none(self):
        df = pd.DataFrame({"a": [np.inf, -np.inf]})
        assert [r["a"] for r in query_engine._records(df)] == [None, None]

    def test_empty_frame(self):
        assert query_engine._records(pd.DataFrame()) == []


class TestClampLimit:
    def test_none_uses_ceiling(self):
        assert query_engine._clamp_limit(None) == query_engine.MAX_ROWS_RETURNED

    def test_over_ceiling_is_capped(self):
        assert query_engine._clamp_limit(10**9) == query_engine.MAX_ROWS_RETURNED

    def test_negative_raises(self):
        # df.head(-5) drops rows from the end instead of erroring.
        with pytest.raises(ValueError):
            query_engine._clamp_limit(-5)

    def test_zero_raises(self):
        with pytest.raises(ValueError):
            query_engine._clamp_limit(0)


# ------------------------------------------------------------ header detect


class TestHeaderDetection:
    def test_skips_title_row(self):
        data = {"values": [["Quarterly Report"], ["Name", "Qty"], ["a", 1]]}
        idx, cols, found = structure.detect_header_row(data)
        assert idx == 1
        assert cols == ["Name", "Qty"]
        assert found

    def test_deduplicates_repeated_headers(self):
        data = {"values": [["Name", "Name", "Qty"]]}
        _, cols, _ = structure.detect_header_row(data)
        assert cols == ["Name", "Name_1", "Qty"]

    def test_blank_cells_get_positional_names(self):
        data = {"values": [["Name", "", "Qty"]]}
        _, cols, _ = structure.detect_header_row(data)
        assert cols == ["Name", "Column_1", "Qty"]

    def test_empty_sheet(self):
        assert structure.detect_header_row({"values": []}) == (0, [], False)

    def test_headerless_data_reports_not_found(self):
        # All-numeric rows: the fallback picks row 0 but flags that nothing
        # actually looked like a header, so scanners can widen the read.
        data = {"values": [[1, 2], [3, 4]]}
        idx, cols, found = structure.detect_header_row(data)
        assert idx == 0
        assert not found


# ----------------------------------------------------------- relationships


def _sheet_with_samples(columns, sampled):
    return {"header_row": 1, "columns": columns, "sampled_values": sampled}


class TestRelationshipInference:
    def test_name_match_plus_value_overlap_infers(self):
        files = {
            "stock.xlsx": {
                "sheets": {
                    "Stock": _sheet_with_samples(
                        ["Material_ID", "Qty"],
                        {"Material_ID": ["TiO2", "CaCO3", "ZnO"]},
                    )
                }
            },
            "prices.xlsx": {
                "sheets": {
                    "Prices": _sheet_with_samples(
                        ["materialid", "Rate"],
                        {"materialid": ["TiO2", "CaCO3", "NaCl"]},
                    )
                }
            },
        }
        rels = structure.infer_relationships(files)
        assert len(rels) == 1
        rel = rels[0]
        assert rel["left"]["column"] == "Material_ID"
        assert rel["right"]["column"] == "materialid"
        assert rel["source"] == "inferred"
        assert rel["confidence"] == pytest.approx(2 / 3, abs=0.01)
        assert rel["evidence"]["name_token"] == "materialid"

    def test_name_match_without_value_overlap_is_not_a_relationship(self):
        files = {
            "a.xlsx": {
                "sheets": {
                    "S": _sheet_with_samples(["Code"], {"Code": ["x1", "x2", "x3"]})
                }
            },
            "b.xlsx": {
                "sheets": {
                    "T": _sheet_with_samples(["Code"], {"Code": ["y1", "y2", "y3"]})
                }
            },
        }
        assert structure.infer_relationships(files) == []

    def test_value_overlap_without_name_match_is_not_a_relationship(self):
        files = {
            "a.xlsx": {
                "sheets": {
                    "S": _sheet_with_samples(
                        ["Material"], {"Material": ["TiO2", "ZnO"]}
                    )
                }
            },
            "b.xlsx": {
                "sheets": {
                    "T": _sheet_with_samples(["Client"], {"Client": ["TiO2", "ZnO"]})
                }
            },
        }
        assert structure.infer_relationships(files) == []


class TestDeclaredRelationships:
    def _write(self, tmp_path, monkeypatch, text):
        path = tmp_path / "relationships.yaml"
        path.write_text(text, encoding="utf-8")
        monkeypatch.setattr(structure, "get_relationships_yaml_path", lambda: path)
        return structure.load_declared_relationships()

    def test_valid_declaration_loads_with_full_confidence(self, tmp_path, monkeypatch):
        rels = self._write(
            tmp_path,
            monkeypatch,
            "relationships:\n"
            "  - left:  {file: a.xlsx, sheet: S, column: Material}\n"
            "    right: {file: b.xlsx, sheet: T, column: Material_ID}\n",
        )
        assert len(rels) == 1
        assert rels[0]["confidence"] == 1.0
        assert rels[0]["source"] == "declared"

    def test_garbage_yaml_degrades_to_empty(self, tmp_path, monkeypatch):
        assert self._write(tmp_path, monkeypatch, "{ not: [valid") == []

    def test_malformed_entries_are_skipped(self, tmp_path, monkeypatch):
        rels = self._write(
            tmp_path,
            monkeypatch,
            "relationships:\n  - left: {file: a.xlsx}\n",
        )
        assert rels == []

    def test_missing_file_is_empty(self, tmp_path, monkeypatch):
        monkeypatch.setattr(
            structure, "get_relationships_yaml_path", lambda: tmp_path / "nope.yaml"
        )
        assert structure.load_declared_relationships() == []


# ------------------------------------------------------------------- joins


def _join_graph(relationships=None):
    return {
        "workspaces": {
            "/ERP": {
                "schema_version": SCHEMA,
                "files": {
                    "stock.xlsx": {
                        "item_id": "stock",
                        "sheets": {
                            "Stock": {"header_row": 1, "columns": ["Material", "Qty"]}
                        },
                    },
                    "prices.xlsx": {
                        "item_id": "prices",
                        "sheets": {
                            "Prices": {
                                "header_row": 1,
                                "columns": ["Material_ID", "Rate"],
                            }
                        },
                    },
                },
                "relationships": relationships or [],
            }
        }
    }


class TestJoinSheets:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        self._graph = _join_graph()
        monkeypatch.setattr(query_engine, "load_graph", lambda: self._graph)
        monkeypatch.setattr(
            query_engine, "workspace_relationships",
            lambda ws: ws.get("relationships", []),
        )

        frames = {
            "stock": pd.DataFrame(
                {"Material": ["TiO2 ", "caco3", "ZnO"], "Qty": [5, 10, 20]}
            ),
            "prices": pd.DataFrame(
                {"Material_ID": ["tio2", "CaCO3", "NaCl"], "Rate": [142.5, 88.0, 3.0]}
            ),
        }

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            return frames[item_id].copy(), None

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _join(self, **kwargs):
        kwargs.setdefault("folder_path", "/ERP")
        kwargs.setdefault("left_file", "stock.xlsx")
        kwargs.setdefault("left_sheet", "Stock")
        kwargs.setdefault("right_file", "prices.xlsx")
        kwargs.setdefault("right_sheet", "Prices")
        return asyncio.run(query_engine.execute_join_sheets(**kwargs))

    def test_explicit_keys_join_with_normalised_matching(self):
        result = self._join(left_on="Material", right_on="Material_ID")
        # 'TiO2 ' joins 'tio2', 'caco3' joins 'CaCO3'; ZnO/NaCl drop (inner).
        assert result["total_matched"] == 2
        rates = sorted(r["Rate"] for r in result["rows"])
        assert rates == [88.0, 142.5]
        assert result["keys"]["source"] == "caller"

    def test_outer_join_keeps_unmatched_rows(self):
        result = self._join(
            left_on="Material", right_on="Material_ID", join_type="outer"
        )
        assert result["total_matched"] == 4  # 2 matches + ZnO + NaCl

    def test_keys_suggested_from_relationship(self):
        self._graph = _join_graph(
            relationships=[
                {
                    "left": {"file": "stock.xlsx", "sheet": "Stock", "column": "Material"},
                    "right": {
                        "file": "prices.xlsx",
                        "sheet": "Prices",
                        "column": "Material_ID",
                    },
                    "confidence": 0.8,
                    "source": "inferred",
                }
            ]
        )
        result = self._join()
        assert result["keys"]["left_on"] == "Material"
        assert result["keys"]["right_on"] == "Material_ID"
        assert result["keys"]["source"] == "inferred"
        assert result["total_matched"] == 2

    def test_refuses_to_guess_without_relationship(self):
        with pytest.raises(ValueError, match="No confident relationship"):
            self._join()

    def test_unknown_join_column_fails_loud(self):
        with pytest.raises(ValueError, match="not found"):
            self._join(left_on="Materiel", right_on="Material_ID")

    def test_unknown_join_type_fails_loud(self):
        with pytest.raises(ValueError, match="join_type"):
            self._join(left_on="Material", right_on="Material_ID", join_type="cross")


class TestDerive:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        graph = {
            "workspaces": {
                "/ERP": {
                    "schema_version": SCHEMA,
                    "files": {
                        "tx.xlsx": {
                            "item_id": "tx",
                            "sheets": {
                                "Transactions": {
                                    "header_row": 1,
                                    "columns": ["Material", "Type", "Qty"],
                                }
                            },
                        }
                    }
                }
            }
        }
        monkeypatch.setattr(query_engine, "load_graph", lambda: graph)

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            return (
                pd.DataFrame(
                    {
                        "Material": ["TiO2", "TiO2", "TiO2", "ZnO"],
                        "Type": ["Receipt", "Consumption", "Receipt", "Receipt"],
                        "Qty": [100, 30, 50, 7],
                    }
                ),
                None,
            )

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _derive(self, components, group_by="Material", **kwargs):
        return asyncio.run(
            query_engine.execute_derive(
                "/ERP", "tx.xlsx", "Transactions", group_by, "Qty",
                components, **kwargs,
            )
        )

    def test_net_stock_in_one_call(self):
        result = self._derive(
            [
                {"conditions": {"Type": "Receipt"}, "sign": 1, "label": "receipts"},
                {"conditions": {"Type": "Consumption"}, "sign": -1, "label": "consumption"},
            ]
        )
        by_material = {r["Material"]: r["net"] for r in result["rows"]}
        assert by_material == {"TiO2": 120.0, "ZnO": 7.0}
        labels = [c["label"] for c in result["components"]]
        assert labels == ["receipts", "consumption"]
        assert result.get("warning") is None

    def test_zero_match_component_is_flagged(self):
        result = self._derive(
            [
                {"conditions": {"Type": "Receipt"}, "sign": 1},
                {"conditions": {"Type": "Returns"}, "sign": -1, "label": "returns"},
            ]
        )
        flagged = [c for c in result["components"] if c["matched_no_rows"]]
        assert [c["label"] for c in flagged] == ["returns"]
        assert "returns" in result["warning"]

    def test_bad_sign_fails_loud(self):
        with pytest.raises(ValueError, match="sign"):
            self._derive([{"conditions": {"Type": "Receipt"}, "sign": 2}])

    def test_unknown_quantity_column_fails_loud(self):
        with pytest.raises(ValueError, match="not found"):
            asyncio.run(
                query_engine.execute_derive(
                    "/ERP", "tx.xlsx", "Transactions", "Material", "Amount",
                    [{"conditions": {"Type": "Receipt"}, "sign": 1}],
                )
            )

    def test_group_by_omitted_returns_single_net_total(self):
        # Receipts across every material (100 + 50 + 7) minus consumption (30).
        result = self._derive(
            [
                {"conditions": {"Type": "Receipt"}, "sign": 1, "label": "receipts"},
                {"conditions": {"Type": "Consumption"}, "sign": -1, "label": "consumption"},
            ],
            group_by=None,
        )
        assert result["rows"] == [{"net": 127.0}]
        assert "Material" not in (result["rows"][0] if result["rows"] else {})

    def test_explicit_empty_group_by_still_errors(self):
        # Only an *omitted* group_by (None) means "no grouping" — an explicit
        # [] is still a caller mistake, not a synonym for the same thing.
        with pytest.raises(ValueError, match="group_by"):
            self._derive(
                [{"conditions": {"Type": "Receipt"}, "sign": 1}], group_by=[]
            )

    def test_base_component_defaults_sign_to_one_and_is_flagged_like_any_other(self):
        result = self._derive(
            [
                {"conditions": {"Type": "Receipt"}, "sign": 1, "label": "receipts"},
                {
                    "conditions": {"Material": "ZnO"},
                    "base": True,
                    "label": "zno_adjustment",
                },
                {
                    "conditions": {"Material": "Nonexistent"},
                    "base": True,
                    "label": "typo_check",
                },
            ]
        )
        by_label = {c["label"]: c for c in result["components"]}
        assert by_label["zno_adjustment"]["sign"] == 1
        assert by_label["zno_adjustment"]["matched_no_rows"] is False
        assert by_label["typo_check"]["matched_no_rows"] is True
        assert "typo_check" in result["warning"]


# --------------------------------------------------------- region scoping


class TestRegionToDfSlice:
    """ranges.region_to_df_slice: the absolute-row <-> DataFrame-index arithmetic
    that filter_sheet/aggregate/derive lean on for `region`. An off-by-one here
    produces a plausible wrong number, not an error, so it is pinned exactly."""

    def test_header_at_top_of_used_range(self):
        # used range starts on the header row itself (row1 == header_abs == 5).
        # DataFrame row 0 is sheet row 6, so body '7:15' starts at df index 1.
        assert ranges.region_to_df_slice("Sheet1!C5:E30", 1, "7:15") == (1, 10)

    def test_used_range_starts_above_the_header(self):
        # used range starts at row 3, header is 3 rows into it (row 5) — same
        # header_abs as above, so the same span must resolve identically.
        assert ranges.region_to_df_slice("Sheet1!C3:E30", 3, "7:15") == (1, 10)

    def test_row_immediately_below_header_is_df_index_zero(self):
        assert ranges.region_to_df_slice("Sheet1!C5:E30", 1, "6:6") == (0, 1)

    def test_second_region_is_offset_correctly(self):
        assert ranges.region_to_df_slice("Sheet1!C5:E30", 1, "20:28") == (14, 23)

    def test_parse_span_tolerates_whitespace(self):
        assert ranges.parse_span(" 7 : 28 ") == (7, 28)

    def test_parse_span_rejects_malformed_text(self):
        with pytest.raises(ValueError):
            ranges.parse_span("not-a-span")


def _region_graph():
    """A sheet with two disjoint table regions (mirrors the Naphthalene/Oleum
    layout) and a single-region sheet, sharing one used-range convention:
    used range starts on the header row, so header_abs == row1 for both."""
    two_table_regions = [
        {
            "type": "table", "range": "5:15", "body": "7:15",
            "header_row": 5, "source": "formula", "confidence": "unconfirmed",
        },
        {
            "type": "table", "range": "18:28", "body": "20:28",
            "header_row": 19, "source": "formula", "confidence": "unconfirmed",
        },
    ]
    return {
        "workspaces": {
            "/ERP": {
                "schema_version": SCHEMA,
                "files": {
                    "stmt.xlsx": {
                        "item_id": "stmt",
                        "sheets": {
                            "Two Tables": {
                                "header_row": 1,
                                "columns": ["Type", "Qty"],
                                "used_range_address": "Two Tables!C5:D30",
                                "regions": two_table_regions,
                            },
                            "One Table": {
                                "header_row": 1,
                                "columns": ["Type", "Qty"],
                                "used_range_address": "One Table!C5:D15",
                                "regions": [two_table_regions[0]],
                            },
                        },
                    }
                },
            }
        }
    }


class TestRegionScoping:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        monkeypatch.setattr(query_engine, "load_graph", lambda: _region_graph())

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            # Two Tables' used range starts at sheet row 5 -> df row 0 == row 6.
            rows = list(range(6, 31)) if sheet_name == "Two Tables" else list(range(6, 16))
            df = pd.DataFrame(
                {
                    "Type": ["Receipt" if r % 2 == 0 else "Return" for r in rows],
                    "Qty": rows,
                    "AbsRow": rows,
                }
            )
            return df, None

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _filter(self, sheet, region=None):
        return asyncio.run(
            query_engine.execute_filter_sheet(
                "stmt.xlsx", sheet, {}, None, None, "/ERP", False, region
            )
        )

    def test_index_and_span_resolve_to_the_same_rows(self):
        by_index = self._filter("Two Tables", region=0)
        by_span = self._filter("Two Tables", region="7:15")
        assert sorted(r["AbsRow"] for r in by_index["rows"]) == list(range(7, 16))
        assert sorted(r["AbsRow"] for r in by_span["rows"]) == list(range(7, 16))
        assert by_index["region_used"] == {"index": 0, "body": "7:15", "row_count": 9}
        assert by_span["region_used"] == by_index["region_used"]

    def test_second_region_is_disjoint_from_the_first(self):
        result = self._filter("Two Tables", region=1)
        assert sorted(r["AbsRow"] for r in result["rows"]) == list(range(20, 29))
        assert result["region_used"]["index"] == 1

    def test_out_of_range_index_errors_listing_available_regions(self):
        with pytest.raises(ValueError, match="out of range"):
            self._filter("Two Tables", region=5)

    def test_unmatched_span_errors_rather_than_falling_back_to_whole_sheet(self):
        with pytest.raises(ValueError, match="matches no region"):
            self._filter("Two Tables", region="100:200")

    def test_omitted_region_on_multi_region_sheet_warns_and_sees_everything(self):
        result = self._filter("Two Tables")
        assert "region_used" not in result
        assert "2 table" in result["region_warning"]
        assert "7:15" in result["region_warning"] and "20:28" in result["region_warning"]
        assert len(result["rows"]) == 25  # unscoped: rows 6..30

    def test_omitted_region_on_single_region_sheet_is_unaffected(self):
        result = self._filter("One Table")
        assert "region_used" not in result
        assert "region_warning" not in result
        assert len(result["rows"]) == 10  # unscoped: rows 6..15

    def test_aggregate_scoped_to_region_sums_only_that_table(self):
        result = asyncio.run(
            query_engine.execute_aggregate(
                "stmt.xlsx", "Two Tables", "Type", "Qty", "sum", None, "/ERP",
                None, 0,
            )
        )
        assert sum(r["Qty"] for r in result["rows"]) == sum(range(7, 16))
        assert result["region_used"]["body"] == "7:15"

    def test_derive_scoped_to_region_matches_that_table_only(self):
        # Region 0 body 7:15: Receipt on even rows (8,10,12,14=44),
        # Return on odd rows (7,9,11,13,15=55) -> net -11.
        result = asyncio.run(
            query_engine.execute_derive(
                "/ERP", "stmt.xlsx", "Two Tables", None, "Qty",
                [
                    {"conditions": {"Type": "Receipt"}, "sign": 1, "label": "receipts"},
                    {"conditions": {"Type": "Return"}, "sign": -1, "label": "returns"},
                ],
                None, 0,
            )
        )
        assert result["rows"] == [{"net": -11.0}]
        assert result["region_used"]["body"] == "7:15"


# --------------------------------------------------- structured conditions


class TestStructuredConditions:
    def _orders(self):
        return pd.DataFrame(
            {
                "Status": ["Closed", "Shipped", "Open", "closed "],
                "Qty": [5, 50, 500, 5000],
                "Batch Date": ["2026-01-15", "2026-02-01", "2026-05-01", ""],
                "Notes": ["ok", "", None, "late"],
            }
        )

    def test_in_list_normalises_like_exact_match(self):
        out = query_engine._apply_conditions(
            self._orders(), {"Status": {"in": ["Closed", "Shipped"]}}
        )
        assert out["Qty"].tolist() == [5, 50, 5000]  # 'closed ' matches too

    def test_between_is_inclusive(self):
        out = query_engine._apply_conditions(
            self._orders(), {"Qty": {"between": [10, 500]}}
        )
        assert out["Qty"].tolist() == [50, 500]

    def test_date_range_combines_two_operators(self):
        out = query_engine._apply_conditions(
            self._orders(),
            {"Batch Date": {">=": "2026-01-01", "<": "2026-04-01"}},
        )
        assert out["Qty"].tolist() == [5, 50]

    def test_is_null_treats_blank_strings_as_null(self):
        out = query_engine._apply_conditions(
            self._orders(), {"Notes": {"is_null": True}}
        )
        assert out["Qty"].tolist() == [50, 500]
        out = query_engine._apply_conditions(
            self._orders(), {"Notes": {"is_null": False}}
        )
        assert out["Qty"].tolist() == [5, 5000]

    def test_flat_and_structured_forms_mix(self):
        out = query_engine._apply_conditions(
            self._orders(),
            {"Status": {"in": ["Closed", "Open"]}, "Qty": ">100"},
        )
        assert out["Qty"].tolist() == [500, 5000]

    def test_unknown_operator_raises(self):
        with pytest.raises(ValueError, match="Unknown operator 'like'"):
            query_engine._apply_conditions(
                self._orders(), {"Status": {"like": "Closed"}}
            )

    def test_malformed_between_raises(self):
        with pytest.raises(ValueError, match="between"):
            query_engine._apply_conditions(
                self._orders(), {"Qty": {"between": [10]}}
            )

    def test_empty_spec_raises(self):
        with pytest.raises(ValueError, match="empty object"):
            query_engine._apply_conditions(self._orders(), {"Qty": {}})

    def test_unknown_column_still_fails_loud(self):
        with pytest.raises(ValueError, match="not found"):
            query_engine._apply_conditions(
                self._orders(), {"Nope": {"in": ["x"]}}
            )


class TestAggregateGrammar:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        graph = {
            "workspaces": {
                "/ERP": {
                    "schema_version": SCHEMA,
                    "files": {
                        "s.xlsx": {
                            "item_id": "id1",
                            "sheets": {
                                "Data": {
                                    "header_row": 1,
                                    "columns": ["Region", "Type", "Revenue"],
                                }
                            },
                        }
                    }
                }
            }
        }
        monkeypatch.setattr(query_engine, "load_graph", lambda: graph)

        async def fake_fetch(item_id, sheet_name, sheet_info=None):
            return (
                pd.DataFrame(
                    {
                        "Region": ["N", "N", "S", "S"],
                        "Type": ["a", "b", "a", "a"],
                        "Revenue": [100, 200, 300, 400],
                    }
                ),
                None,
            )

        monkeypatch.setattr(query_engine, "_fetch_sheet_data", fake_fetch)

    def _agg(self, group_by, having=None):
        return asyncio.run(
            query_engine.execute_aggregate(
                "s.xlsx", "Data", group_by, "Revenue", "sum", None, "/ERP", having
            )
        )

    def test_multi_column_group_by(self):
        rows = self._agg(["Region", "Type"])["rows"]
        assert {"Region": "S", "Type": "a", "Revenue": 700.0} in [
            {k: v for k, v in r.items()} for r in rows
        ]
        assert len(rows) == 3

    def test_having_filters_aggregated_rows(self):
        rows = self._agg("Region", having={"Revenue": ">400"})["rows"]
        assert len(rows) == 1
        assert rows[0]["Region"] == "S"

    def test_having_unknown_column_raises(self):
        with pytest.raises(ValueError, match="not found"):
            self._agg("Region", having={"Nope": ">1"})

    def test_unknown_group_column_raises(self):
        with pytest.raises(ValueError, match="group_by column"):
            self._agg(["Region", "Nope"])


# --------------------------------------------------------- structure drift


class TestStructureDrift:
    def test_matching_header_reports_no_drift(self):
        report = query_engine._drift_report(
            ["Name", "Qty"], {"columns": ["Name", "Qty"]}
        )
        assert report is None

    def test_changed_header_reports_both_versions(self):
        report = query_engine._drift_report(
            ["Inserted", "Name", "Qty"], {"columns": ["Name", "Qty"]}
        )
        assert report["structure_drift"] is True
        assert report["stored_header"] == ["Name", "Qty"]
        assert report["live_header"] == ["Inserted", "Name", "Qty"]
        assert "scan_workspace" in report["recommendation"]

    def test_unscanned_sheet_cannot_drift(self):
        assert query_engine._drift_report(["A"], {}) is None
        assert query_engine._drift_report(["A"], None) is None

    def test_fetch_flags_drift_but_still_answers(self, monkeypatch):
        async def fake_used_range(item_id, sheet_name, *args, **kwargs):
            return {"values": [["Renamed", "Qty"], ["a", 1]]}

        monkeypatch.setattr(query_engine, "get_used_range", fake_used_range)
        df, drift = asyncio.run(
            query_engine._fetch_sheet_data(
                "i", "S", {"header_row": 1, "columns": ["Name", "Qty"]}
            )
        )
        assert drift["structure_drift"] is True
        assert list(df.columns) == ["Renamed", "Qty"]  # live header wins
        assert len(df) == 1

    def test_fetch_uses_scan_style_column_names(self, monkeypatch):
        # Blank headers must get the same positional names the graph shows
        # ('Column_1'), not a different scheme the agent has never seen.
        async def fake_used_range(item_id, sheet_name, *args, **kwargs):
            return {"values": [["Name", "", "Qty"], ["a", "x", 1]]}

        monkeypatch.setattr(query_engine, "get_used_range", fake_used_range)
        df, _ = asyncio.run(query_engine._fetch_sheet_data("i", "S", {"header_row": 1}))
        assert list(df.columns) == ["Name", "Column_1", "Qty"]


# ------------------------------------------------------------ serial dates


class TestDateFormatDetection:
    @pytest.mark.parametrize(
        "fmt",
        ["m/d/yyyy", "dd-mmm-yy", "yyyy-mm-dd", "d-mmm", "hh:mm:ss", "[$-409]d-mmm;@"],
    )
    def test_date_formats_detected(self, fmt):
        assert structure.is_date_number_format(fmt)

    @pytest.mark.parametrize(
        "fmt", ["General", "0.00", "#,##0.00", "$#,##0.00", "@", "0%", None, ""]
    )
    def test_non_date_formats_rejected(self, fmt):
        assert not structure.is_date_number_format(fmt)

    def test_quoted_literals_do_not_confuse_detection(self):
        # 'd' inside a quoted literal is text, not a day token.
        assert not structure.is_date_number_format('"days" 0')

    def test_detect_date_columns_uses_first_data_cell(self):
        values = [["Name", "When", "Qty"], ["a", 45000, 5], ["b", 45001, 6]]
        formats = [["General"] * 3, ["General", "m/d/yyyy", "0"], ["General", "m/d/yyyy", "0"]]
        types = structure.detect_date_columns(values, formats, 0, ["Name", "When", "Qty"])
        assert types == {"When": "date"}

    def test_blank_leading_cells_are_skipped(self):
        values = [["Name", "When"], ["a", None], ["b", 45000]]
        formats = [["General"] * 2, ["General", "General"], ["General", "yyyy-mm-dd"]]
        types = structure.detect_date_columns(values, formats, 0, ["Name", "When"])
        assert types == {"When": "date"}


class TestSerialConversion:
    def test_whole_serial_becomes_date(self):
        assert query_engine._serial_to_iso(45000) == "2023-03-15"

    def test_fractional_serial_keeps_time(self):
        assert query_engine._serial_to_iso(45000.5) == "2023-03-15T12:00:00"

    def test_epoch_offset_is_lotus_compatible(self):
        # Serial 1 is 1900-01-01 in the 1900 date system.
        assert query_engine._serial_to_iso(1) == "1899-12-31"
        assert query_engine._serial_to_iso(2) == "1900-01-01"

    def test_non_numeric_passes_through(self):
        assert query_engine._serial_to_iso("already a string") == "already a string"
        assert query_engine._serial_to_iso(None) is None

    def test_absurd_serial_is_left_alone(self):
        assert query_engine._serial_to_iso(1e12) == 1e12

    def test_convert_only_flagged_columns(self):
        df = pd.DataFrame({"When": [45000, "x"], "Qty": [45000, 1]})
        out = query_engine._convert_date_columns(df, {"When": "date"})
        assert out["When"].tolist() == ["2023-03-15", "x"]
        assert out["Qty"].tolist() == [45000, 1]  # untouched


class TestDateConditions:
    def _dates(self):
        return pd.DataFrame(
            {
                "Batch Date": ["2026-01-15", "2026-03-01", "2025-12-31", "", "n/a"],
                "Qty": [1, 2, 3, 4, 5],
            }
        )

    def test_iso_range_conditions(self):
        out = query_engine._apply_conditions(
            self._dates(), {"Batch Date": ">=2026-01-01"}
        )
        assert out["Qty"].tolist() == [1, 2]
        out = query_engine._apply_conditions(
            self._dates(), {"Batch Date": "<2026-01-01"}
        )
        assert out["Qty"].tolist() == [3]

    def test_non_date_rows_never_match_date_comparisons(self):
        out = query_engine._apply_conditions(
            self._dates(), {"Batch Date": ">=1000-01-01"}
        )
        assert out["Qty"].tolist() == [1, 2, 3]  # blanks and 'n/a' excluded

    def test_datetime_rows_compare_correctly_against_date_bounds(self):
        df = pd.DataFrame({"When": ["2026-01-01T09:30:00"], "Qty": [1]})
        assert len(query_engine._apply_conditions(df, {"When": ">=2026-01-01"})) == 1
        assert len(query_engine._apply_conditions(df, {"When": "<2026-01-02"})) == 1

    def test_garbage_bound_still_raises(self):
        with pytest.raises(ValueError, match="neither a number nor an ISO date"):
            query_engine._apply_conditions(self._dates(), {"Batch Date": ">soon"})


# ------------------------------------------------------------- A1 notation


class TestRanges:
    def test_col_roundtrip(self):
        from excelmcp import ranges

        for index, letters in [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (703, "AAA")]:
            assert ranges.index_to_col(index) == letters
            assert ranges.col_to_index(letters) == index

    def test_parse_strips_sheet_prefix(self):
        from excelmcp import ranges

        assert ranges.parse_range("Sheet1!B3:H500") == (2, 3, 8, 500)
        assert ranges.parse_range("'My Sheet'!A1") == (1, 1, 1, 1)

    def test_build_range_and_cell(self):
        from excelmcp import ranges

        assert ranges.build_range(2, 3, 8, 500) == "B3:H500"
        assert ranges.build_cell(8, 347) == "H347"

    def test_garbage_raises(self):
        from excelmcp import ranges

        with pytest.raises(ValueError):
            ranges.parse_range("not-an-address")


# ------------------------------------------------------- bounded scan reads


class TestScanSheetStructure:
    def _fakes(self, monkeypatch, meta, windows, full=None, formulas=None):
        calls = []

        async def fake_used_range(item_id, sheet, session=None, *, select=None, **kw):
            if select and "formulas" in select:
                calls.append(("formulas", select))
                return {"formulas": formulas or []}
            if select and "values" not in select:
                calls.append(("meta", select))
                return meta
            calls.append(("full", select))
            return full or {"values": []}

        async def fake_get_range(item_id, sheet, address, session=None, **kw):
            calls.append(("range", address))
            return windows.get(address, {"values": []})

        monkeypatch.setattr(structure, "get_used_range", fake_used_range)
        monkeypatch.setattr(structure, "get_range", fake_get_range)
        return calls

    def test_header_in_window_avoids_full_read(self, monkeypatch):
        calls = self._fakes(
            monkeypatch,
            meta={"address": "Sheet1!A1:C500", "rowCount": 500, "columnCount": 3},
            windows={
                "A1:C10": {"values": [["Name", "Qty", "Status"], ["a", 1, "x"]]},
                "A2:C500": {"values": [["a", 1, "Open"], ["b", 2, "Closed"]]},
            },
        )
        entry = asyncio.run(structure._scan_sheet_structure("i", "S", None))
        assert entry["columns"] == ["Name", "Qty", "Status"]
        assert entry["header_row"] == 1
        assert entry["used_range_address"] == "Sheet1!A1:C500"
        assert entry["row_count"] == 500
        assert entry["approx_row_count"] == 499
        assert entry["sampled_values"]["Status"] == ["Closed", "Open"]
        # Metadata, formulas, header window, sampling window — never the whole sheet.
        assert [kind for kind, _ in calls] == ["meta", "formulas", "range", "range"]

    def test_no_header_in_window_falls_back_to_full_read(self, monkeypatch):
        numeric = {"values": [[1, 2]] * 10}
        calls = self._fakes(
            monkeypatch,
            meta={"address": "Sheet1!A1:B40", "rowCount": 40, "columnCount": 2},
            windows={"A1:B10": numeric},
            full={"values": [[1, 2]] * 11 + [["Name", "Qty"], ["a", 1]]},
        )
        entry = asyncio.run(structure._scan_sheet_structure("i", "S", None))
        assert entry["columns"] == ["Name", "Qty"]
        assert entry["header_row"] == 12
        # Falls back to the full read, then still samples values for routing.
        assert [kind for kind, _ in calls] == [
            "meta",
            "formulas",
            "range",
            "full",
            "range",
        ]

    def test_empty_sheet_returns_none(self, monkeypatch):
        self._fakes(
            monkeypatch,
            meta={"address": "", "rowCount": 0, "columnCount": 0},
            windows={},
        )
        assert asyncio.run(structure._scan_sheet_structure("i", "S", None)) is None

    def test_wide_sheet_refetches_header_row_at_full_width(self, monkeypatch):
        wide_header = [f"C{i}" for i in range(60)]
        calls = self._fakes(
            monkeypatch,
            meta={"address": "Sheet1!A1:BH100", "rowCount": 100, "columnCount": 60},
            windows={
                "A1:AZ10": {"values": [wide_header[:52], ["x"] * 52]},
                "A1:BH1": {"values": [wide_header]},
            },
        )
        entry = asyncio.run(structure._scan_sheet_structure("i", "S", None))
        assert len(entry["columns"]) == 60
        assert entry["columns"][-1] == "C59"
        assert ("range", "A1:BH1") in calls


# ---------------------------------------------------------------- storage


class TestStorage:
    def test_atomic_write_replaces_content(self, tmp_path):
        target = tmp_path / "f.json"
        storage.atomic_write_text(target, "one")
        storage.atomic_write_text(target, "two")
        assert target.read_text(encoding="utf-8") == "two"

    def test_no_temp_files_left_behind(self, tmp_path):
        target = tmp_path / "f.json"
        storage.atomic_write_text(target, "x")
        assert [p.name for p in tmp_path.iterdir()] == ["f.json"]

    def test_read_json_tolerates_corruption(self, tmp_path):
        bad = tmp_path / "graph.json"
        bad.write_text("{ this is not json", encoding="utf-8")
        # A truncated graph.json used to raise out of every tool call.
        assert storage.read_json(bad, {"workspaces": {}}) == {"workspaces": {}}

    def test_read_json_missing_file(self, tmp_path):
        assert storage.read_json(tmp_path / "nope.json", []) == []

    def test_load_graph_survives_corrupt_file(self, tmp_path, monkeypatch):
        path = tmp_path / "graph.json"
        path.write_text("garbage", encoding="utf-8")
        monkeypatch.setattr(structure, "get_graph_path", lambda: path)
        assert structure.load_graph() == {"workspaces": {}}

    def test_load_graph_rejects_wrong_shape(self, tmp_path, monkeypatch):
        path = tmp_path / "graph.json"
        path.write_text('["not", "a", "dict"]', encoding="utf-8")
        monkeypatch.setattr(structure, "get_graph_path", lambda: path)
        assert structure.load_graph() == {"workspaces": {}}


# --------------------------------------------------------------- embeddings


@pytest.fixture
def fake_index(tmp_path, monkeypatch):
    """Redirects the index to tmp_path and stubs the embedding model."""
    monkeypatch.setattr(embeddings, "VECTORS_PATH", tmp_path / "vectors.npy")
    monkeypatch.setattr(embeddings, "METADATA_PATH", tmp_path / "metadata.json")
    monkeypatch.setattr(embeddings, "_LEGACY_INDEX_PATH", tmp_path / "index.bin")
    monkeypatch.setattr(embeddings, "get_config_dir", lambda: tmp_path)
    monkeypatch.setattr(embeddings, "_vectors", None)
    monkeypatch.setattr(embeddings, "_metadata", None)

    def fake_embed(texts):
        # Deterministic pseudo-embedding: bucket by first character.
        out = []
        for t in texts:
            vec = [0.0] * embeddings.DIM
            vec[ord(t[0]) % embeddings.DIM] = 1.0
            out.append(vec)
        return out

    monkeypatch.setattr(embeddings, "embed", fake_embed)
    return tmp_path


def _files(prefix):
    return {
        f"{prefix}.xlsx": {
            "sheets": {
                "Sheet1": {"description": f"{prefix} sheet one"},
            }
        }
    }


class TestEmbeddings:
    def test_roundtrip(self, fake_index):
        embeddings.update_embeddings("/ERP", _files("alpha"))
        results = embeddings.search("/ERP", "alpha sheet one", n_results=3)
        assert len(results) == 1
        assert results[0]["file"] == "alpha.xlsx"

    def test_second_workspace_does_not_wipe_the_first(self, fake_index):
        # update_embeddings used to rebuild the whole index from one workspace,
        # silently deleting every other workspace's embeddings.
        embeddings.update_embeddings("/ERP", _files("alpha"))
        embeddings.update_embeddings("/Finance", _files("beta"))
        assert embeddings.collection_count("/ERP") == 1
        assert embeddings.collection_count("/Finance") == 1

    def test_search_is_scoped_to_its_workspace(self, fake_index):
        # The workspace filter used to run *after* top-k, so a query whose
        # nearest neighbours all belonged to another workspace returned nothing.
        embeddings.update_embeddings("/ERP", _files("alpha"))
        embeddings.update_embeddings("/Finance", _files("beta"))
        results = embeddings.search("/Finance", "beta sheet one", n_results=5)
        assert results
        assert {r["workspace"] for r in results} == {"/Finance"}

    def test_unknown_workspace_returns_empty(self, fake_index):
        embeddings.update_embeddings("/ERP", _files("alpha"))
        assert embeddings.search("/Nope", "anything") == []

    def test_trailing_slash_is_normalised(self, fake_index):
        embeddings.update_embeddings("/ERP/", _files("alpha"))
        assert embeddings.collection_count("/ERP") == 1

    def test_desynced_index_is_rejected_not_mistrusted(self, fake_index):
        embeddings.update_embeddings("/ERP", _files("alpha"))
        # Simulate a crash between the metadata and vector writes.
        embeddings.METADATA_PATH.write_text(
            json.dumps(
                [
                    {"workspace": "/ERP", "file": "a", "sheet": "s", "description": "d"},
                    {"workspace": "/ERP", "file": "b", "sheet": "s", "description": "d"},
                ]
            ),
            encoding="utf-8",
        )
        embeddings._vectors = None
        embeddings._metadata = None
        assert embeddings.search("/ERP", "anything") == []

    def test_empty_descriptions_raise(self, fake_index):
        with pytest.raises(RuntimeError, match="No sheet descriptions"):
            embeddings.update_embeddings("/ERP", {"x.xlsx": {"sheets": {}}})


# ------------------------------------------------------- value-level routing


class TestValueSampling:
    def test_distinct_strings_are_collected_sorted(self):
        rows = [["b", 1], ["a", 2], ["b", 3]]
        sampled = structure.sample_column_values(rows, ["Client", "Qty"])
        assert sampled == {"Client": ["a", "b"]}  # Qty is numeric — excluded

    def test_high_cardinality_columns_are_dropped(self):
        rows = [[f"value-{i}"] for i in range(60)]
        assert structure.sample_column_values(rows, ["FreeText"]) == {}

    def test_long_values_and_blanks_are_skipped(self):
        rows = [["x" * 100], [" "], ["ok"]]
        assert structure.sample_column_values(rows, ["Notes"]) == {"Notes": ["ok"]}

    def test_whitespace_is_collapsed(self):
        rows = [["BESTEX  Ltd"], ["BESTEX Ltd"]]
        sampled = structure.sample_column_values(rows, ["Client"])
        assert sampled == {"Client": ["BESTEX Ltd"]}

    def test_description_folds_in_sampled_values(self):
        desc = structure.generate_sheet_description(
            "f.xlsx", "Rates", ["Client", "Rate"], {"Client": ["BESTEX", "Veda"]}
        )
        assert "Client values include: BESTEX, Veda." in desc


class TestLexicalRerank:
    def _identical_schema_files(self):
        # Same description text → identical fake embedding vectors. Only the
        # sampled values differ, exactly like twenty contract workbooks
        # sharing one schema.
        def contract(client):
            return {
                "sheets": {
                    "Rates": {
                        "description": "same schema either way",
                        "columns": ["Client", "Rate"],
                        "sampled_values": {"Client": [client]},
                    }
                }
            }

        return {
            "CON001_BESTEX.xlsx": contract("BESTEX"),
            "CON002_GREEN.xlsx": contract("Greenfield Coatings"),
        }

    def test_rerank_separates_identical_schemas(self, fake_index):
        embeddings.update_embeddings("/ERP", self._identical_schema_files())
        results = embeddings.search("/ERP", "contracted rate for BESTEX", n_results=2)
        assert results[0]["file"] == "CON001_BESTEX.xlsx"
        assert results[0]["lexical_overlap"] > results[1]["lexical_overlap"]
        assert results[0]["score"] > results[1]["score"]

    def test_scores_expose_both_stages(self, fake_index):
        embeddings.update_embeddings("/ERP", self._identical_schema_files())
        (top, _) = embeddings.search("/ERP", "rate for BESTEX", n_results=2)
        assert set(top) >= {"score", "vector_score", "lexical_overlap"}


class TestRoutingSummary:
    def test_near_ties_are_flagged_ambiguous(self):
        summary = query_engine._routing_summary(
            [
                {"file": "a.xlsx", "sheet": "S", "score": 0.90},
                {"file": "b.xlsx", "sheet": "S", "score": 0.895},
                {"file": "c.xlsx", "sheet": "S", "score": 0.50},
            ]
        )
        assert summary["routing_ambiguous"] is True
        assert [t["file"] for t in summary["near_ties"]] == ["a.xlsx", "b.xlsx"]

    def test_clear_winner_is_not_ambiguous(self):
        summary = query_engine._routing_summary(
            [
                {"file": "a.xlsx", "sheet": "S", "score": 0.90},
                {"file": "b.xlsx", "sheet": "S", "score": 0.60},
            ]
        )
        assert summary["routing_ambiguous"] is False
        assert "near_ties" not in summary

    def test_bad_n_results_raises(self):
        with pytest.raises(ValueError, match="n_results"):
            asyncio.run(query_engine.execute_query("q", "/ERP", n_results=0))


# ------------------------------------------------------- device-code recovery


class _ScriptedApp:
    """A PublicClientApplication stand-in returning scripted token-endpoint results."""

    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.codes_issued = 0

    def initiate_device_flow(self, scopes=None):
        self.codes_issued += 1
        return {
            "user_code": f"CODE{self.codes_issued}",
            "verification_uri": "https://microsoft.com/devicelogin",
        }

    def acquire_token_by_device_flow(self, flow):
        return self.outcomes.pop(0)


# The exact shape Entra returns when our own poll redeems a code twice: the
# first response carried the tokens but never arrived, so the retry found the
# code spent. The user sees "All done!" in the browser and a failure in the CLI.
_SPENT = {
    "error": "invalid_grant",
    "error_description": (
        "AADSTS70000: The provided value for the input parameter 'device_code' "
        "has already been used. Trace ID: t Correlation ID: c"
    ),
}
_EXPIRED = {
    "error": "expired_token",
    "error_description": "AADSTS70019: Verification code expired.",
}
_DECLINED = {"error": "authorization_declined", "error_description": "User declined."}
_GOOD = {"access_token": "tok", "expires_in": 3600}


class TestDeviceFlowRecovery:
    def _run(self, outcomes):
        app = _ScriptedApp(outcomes)
        announced = []
        cache = msal.SerializableTokenCache()
        result = asyncio.run(
            auth._run_device_flow(app, cache, lambda f: announced.append(f["user_code"]))
        )
        return app, announced, result

    def test_spent_code_starts_a_fresh_flow(self):
        app, announced, result = self._run([_SPENT, _GOOD])
        assert result["access_token"] == "tok"
        assert app.codes_issued == 2
        # A retry is useless unless the user is shown the new code.
        assert announced == ["CODE1", "CODE2"]

    def test_expired_code_starts_a_fresh_flow(self):
        app, announced, result = self._run([_EXPIRED, _GOOD])
        assert result["access_token"] == "tok"
        assert announced == ["CODE1", "CODE2"]

    def test_success_on_the_first_try_issues_one_code(self):
        app, announced, result = self._run([_GOOD])
        assert app.codes_issued == 1
        assert announced == ["CODE1"]

    def test_retries_are_bounded(self):
        app = _ScriptedApp([_SPENT] * auth._DEVICE_FLOW_ATTEMPTS)
        with pytest.raises(RuntimeError, match="never reached this machine"):
            asyncio.run(auth._run_device_flow(app, msal.SerializableTokenCache(), lambda f: None))
        assert app.codes_issued == auth._DEVICE_FLOW_ATTEMPTS

    def test_declined_is_not_retried(self):
        # Retrying a deliberate refusal just re-prompts a user who said no.
        app = _ScriptedApp([_DECLINED, _GOOD])
        with pytest.raises(RuntimeError, match="declined in the browser"):
            asyncio.run(auth._run_device_flow(app, msal.SerializableTokenCache(), lambda f: None))
        assert app.codes_issued == 1

    def test_flow_that_never_starts_reports_the_reason(self):
        class Broken:
            def initiate_device_flow(self, scopes=None):
                return {"error": "invalid_client", "error_description": "bad client id"}

        with pytest.raises(RuntimeError, match="bad client id"):
            asyncio.run(
                auth._run_device_flow(Broken(), msal.SerializableTokenCache(), lambda f: None)
            )

    def test_consent_failure_explains_the_permission(self):
        assert "Files.Read.All" in auth._device_flow_message(
            {"error": "invalid_grant", "error_description": "AADSTS65004: user declined consent"}
        )
