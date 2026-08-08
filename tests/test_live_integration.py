"""Live integration tests against a real OneDrive workspace.

Opt-in: these are skipped unless EXCELMCP_TEST_FOLDER is set, because they need
a completed `excelmcp-setup`, network access, and real Microsoft credentials.

    EXCELMCP_TEST_FOLDER=/ERP pytest tests/test_live_integration.py -v

This replaces the previous test_suite.py, which had drifted badly: it called
execute_cross_file_aggregate with a positional `sources` list the function had
never accepted, and json.loads()'d results that are dicts, so every assertion
either errored or was vacuous.
"""

import os

import pytest

from excelmcp.embeddings import search
from excelmcp.query_engine import (
    execute_aggregate,
    execute_cross_file_aggregate,
    execute_filter_sheet,
    execute_inspect_file,
    execute_query,
)
from excelmcp.structure import load_graph

FOLDER = os.environ.get("EXCELMCP_TEST_FOLDER", "")

pytestmark = pytest.mark.skipif(
    not FOLDER, reason="Set EXCELMCP_TEST_FOLDER=/YourFolder to run live tests."
)


@pytest.fixture(scope="module")
def workspace():
    ws = load_graph().get("workspaces", {}).get(FOLDER.rstrip("/"))
    if not ws or not ws.get("files"):
        pytest.skip(f"No scanned workspace at '{FOLDER}'. Run excelmcp-setup first.")
    return ws


@pytest.fixture(scope="module")
def first_sheet(workspace):
    for file_name, info in workspace["files"].items():
        for sheet_name, sheet in info.get("sheets", {}).items():
            if sheet.get("columns"):
                return file_name, sheet_name, sheet["columns"]
    pytest.skip("No sheet with columns in the scanned workspace.")


class TestStructureInvariant:
    def test_graph_stores_no_cell_values(self, workspace):
        """The core design rule: structure is cached, data never is."""
        for file_name, info in workspace["files"].items():
            for sheet_name, sheet in info.get("sheets", {}).items():
                assert "values" not in sheet, f"cell values cached in {file_name}/{sheet_name}"
                assert "rows" not in sheet, f"cell values cached in {file_name}/{sheet_name}"
                assert set(sheet) <= {
                    "header_row",
                    "columns",
                    "description",
                    "approx_row_count",
                }

    def test_every_file_has_an_item_id(self, workspace):
        for file_name, info in workspace["files"].items():
            assert info.get("item_id"), f"{file_name} has no item_id"


class TestInspect:
    async def test_returns_structure(self, first_sheet):
        file_name, sheet_name, _ = first_sheet
        result = await execute_inspect_file(file_name, FOLDER)
        assert result["file"] == file_name
        assert sheet_name in result["sheets"]

    async def test_unknown_file_lists_alternatives(self):
        result = await execute_inspect_file("definitely-not-real.xlsx", FOLDER)
        assert "error" in result
        assert "available" in result


class TestFilter:
    async def test_fetches_live_rows(self, first_sheet):
        file_name, sheet_name, _ = first_sheet
        result = await execute_filter_sheet(file_name, sheet_name, {}, None, 10, FOLDER)
        assert isinstance(result["rows"], list)
        assert len(result["rows"]) <= 10
        assert result["total_matched"] >= len(result["rows"])

    async def test_rows_are_json_serialisable(self, first_sheet):
        import json

        file_name, sheet_name, _ = first_sheet
        result = await execute_filter_sheet(file_name, sheet_name, {}, None, 50, FOLDER)
        json.dumps(result["rows"])  # NaN or numpy scalars would raise here

    async def test_unknown_sheet_names_the_valid_ones(self, first_sheet):
        file_name, _, _ = first_sheet
        with pytest.raises(ValueError, match="Available sheets"):
            await execute_filter_sheet(file_name, "NoSuchSheet", {}, None, 5, FOLDER)


class TestAggregate:
    async def test_count_groups(self, first_sheet):
        file_name, sheet_name, columns = first_sheet
        if len(columns) < 2:
            pytest.skip("needs at least two columns")
        result = await execute_aggregate(
            file_name, sheet_name, columns[0], columns[1], "count", None, FOLDER
        )
        assert isinstance(result["rows"], list)

    async def test_rejects_unknown_operation(self, first_sheet):
        file_name, sheet_name, columns = first_sheet
        with pytest.raises(ValueError, match="Unknown operation"):
            await execute_aggregate(
                file_name, sheet_name, columns[0], columns[0], "median", None, FOLDER
            )


class TestCrossFile:
    async def test_aggregates_across_files(self, workspace, first_sheet):
        _, sheet_name, columns = first_sheet
        result = await execute_cross_file_aggregate(
            FOLDER, sheet_name, columns[-1], "count", None
        )
        assert result["operation"] == "count"
        assert isinstance(result["per_file"], dict)

    async def test_unknown_sheet_raises(self):
        with pytest.raises(ValueError, match="No files found"):
            await execute_cross_file_aggregate(FOLDER, "NoSuchSheet", "x", "sum", None)


class TestQuery:
    def test_search_returns_this_workspace_only(self):
        results = search(FOLDER, "data", n_results=5)
        assert all(r["workspace"] == FOLDER.rstrip("/") for r in results)

    async def test_query_returns_live_results(self):
        result = await execute_query("show me all data", FOLDER)
        assert "results" in result
        assert "Run scan_workspace" not in str(result.get("message", ""))
