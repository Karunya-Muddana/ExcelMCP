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
                "Sheet1": {"description": f"{prefix} sheet one", "use_for": []},
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
