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

from excelmcp import auth, embeddings, graph_client, query_engine, storage, structure
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
        with pytest.raises(ValueError, match="not a number"):
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
    return {"workspaces": {"/ERP": {"files": files}}}


class TestCrossFileFold:
    """The incremental fold must agree with what pd.concat used to produce."""

    def _patch(self, monkeypatch, frames_by_item, n_files):
        monkeypatch.setattr(
            query_engine, "load_graph", lambda: _fold_workspace(n_files)
        )

        async def fake_fetch(item_id, sheet_name, header_row=1, **kwargs):
            frame = frames_by_item[item_id]
            if isinstance(frame, Exception):
                raise frame
            return frame.copy()

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
    return {"workspaces": {"/ERP": {"files": files}}}


class TestCrossFileCompleteness:
    @pytest.fixture(autouse=True)
    def patched(self, monkeypatch):
        monkeypatch.setattr(query_engine, "load_graph", _variant_workspace)

        async def fake_fetch(item_id, sheet_name, header_row=1, **kwargs):
            return pd.DataFrame({"Client": ["A", "B"], "Amount": [100, 100]})

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
        idx, cols = structure.detect_header_row(data)
        assert idx == 1
        assert cols == ["Name", "Qty"]

    def test_deduplicates_repeated_headers(self):
        data = {"values": [["Name", "Name", "Qty"]]}
        _, cols = structure.detect_header_row(data)
        assert cols == ["Name", "Name_1", "Qty"]

    def test_blank_cells_get_positional_names(self):
        data = {"values": [["Name", "", "Qty"]]}
        _, cols = structure.detect_header_row(data)
        assert cols == ["Name", "Column_1", "Qty"]

    def test_empty_sheet(self):
        assert structure.detect_header_row({"values": []}) == (0, [])


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
