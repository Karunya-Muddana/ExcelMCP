"""Single-cell lookup: one call in, one provenanced value out.

Two entry points back the two new tools:

* ``execute_get_cell`` — deterministic, addressed read of one cell.
* ``execute_lookup`` — semantic one-shot lookup: resolve a (key column, key
  value, return column) triple from a natural-language query or take it
  explicitly, route to candidate sheets using the value-level index built at
  scan time, then read **only** the key column and the matched row — two small
  reads instead of a whole-sheet fetch.

Everything here is deterministic and dependency-light: routing uses the graph
plus sampled values and lexical matching, never a nested LLM call. Every
failure mode is explicit — a wrong number from this module looks maximally
authoritative, so no response is ever a bare value without provenance or an
ambiguity reason.
"""

import asyncio
import difflib
import re
from typing import Any, Optional

from excelmcp.embeddings import search
from excelmcp.graph_client import get_named_range, get_range
from excelmcp.query_engine import _serial_to_iso
from excelmcp.ranges import build_cell, build_range, parse_range
from excelmcp.storage import log
from excelmcp.structure import (
    is_date_number_format,
    load_graph,
    normalise_name,
    require_current_schema,
)

_WORD_RE = re.compile(r"[a-z0-9]+")

_MAX_CANDIDATE_SHEETS = 3
_MAX_KEY_PROBES_PER_SHEET = 2
_MAX_AMBIGUOUS_ROWS = 3
_MAX_EXPLICIT_CANDIDATES = 10
_RETURN_COLUMN_MIN_OVERLAP = 0.5
_FUZZY_SUGGESTION_CUTOFF = 0.75


def _tokens(text: str) -> set[str]:
    return set(_WORD_RE.findall(str(text).lower()))


def _norm(value: Any) -> str:
    return normalise_name(value)


def _key_match(cell: Any, key_value: Any) -> bool:
    """Normalised string equality, with a numeric fallback so 45 matches 45.0."""
    if _norm(cell) == _norm(key_value):
        return True
    try:
        return float(cell) == float(key_value)
    except (TypeError, ValueError):
        return False


def _values_agree(a: Any, b: Any) -> bool:
    try:
        fa, fb = float(a), float(b)
        return abs(fa - fb) <= 1e-9 * max(1.0, abs(fa), abs(fb))
    except (TypeError, ValueError):
        return _norm(a) == _norm(b)


def _resolved_type(value: Any, was_date: bool) -> str:
    if was_date:
        return "date"
    if value is None or (isinstance(value, str) and not value.strip()):
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    return "string"


# --------------------------------------------------------------- get_cell


def _workspace_files(folder_path: str) -> dict[str, Any]:
    workspace = load_graph().get("workspaces", {}).get(folder_path)
    if not workspace:
        raise ValueError(
            f"Workspace '{folder_path}' not found. Run scan_workspace first."
        )
    require_current_schema(workspace, folder_path)
    return workspace.get("files", {})


def _file_entry(files: dict[str, Any], file_name: str) -> dict[str, Any]:
    file_info = files.get(file_name)
    if not file_info:
        raise ValueError(
            f"File '{file_name}' not found. Available: {list(files.keys())}"
        )
    return file_info


async def execute_get_cell(
    file_name: str, sheet: str, address: str, folder_path: str
) -> dict[str, Any]:
    """One addressed cell, one Graph request, dates already converted."""
    folder_path = folder_path.rstrip("/")
    file_info = _file_entry(_workspace_files(folder_path), file_name)
    if sheet not in file_info.get("sheets", {}):
        raise ValueError(
            f"Sheet '{sheet}' not found in '{file_name}'. "
            f"Available sheets: {list(file_info.get('sheets', {}).keys())}"
        )
    item_id = file_info["item_id"]

    try:
        col1, row1, col2, row2 = parse_range(address)
    except ValueError:
        # Not A1 notation — try it as a workbook-scoped named range.
        data = await get_named_range(
            item_id, address, select="values,address,numberFormat"
        )
        resolved = str(data.get("address") or "")
        col1, row1, col2, row2 = parse_range(resolved)
        if (col1, row1) != (col2, row2):
            raise ValueError(
                f"Named range '{address}' resolves to {resolved}, which is not "
                f"a single cell. get_cell reads exactly one cell."
            )
    else:
        if (col1, row1) != (col2, row2):
            raise ValueError(
                f"get_cell reads a single cell; '{address}' is a range. "
                f"Use filter_sheet for multi-cell reads."
            )
        data = await get_range(
            item_id, sheet, address, select="values,address,numberFormat"
        )

    values = data.get("values") or [[None]]
    formats = data.get("numberFormat") or [[None]]
    raw = values[0][0] if values and values[0] else None
    fmt = formats[0][0] if formats and formats[0] else None
    was_date = is_date_number_format(fmt)
    value = _serial_to_iso(raw) if was_date else raw

    resolved_address = str(data.get("address") or address).split("!")[-1]
    return {
        "value": value,
        "address": resolved_address,
        "file": file_name,
        "sheet": sheet,
        "resolved_type": _resolved_type(value, was_date),
    }


# ----------------------------------------------------------------- lookup


def _scope_allows(scope: Optional[dict], file_name: str, sheet_name: str) -> bool:
    if not scope:
        return True
    want_file = scope.get("file")
    want_sheet = scope.get("sheet")
    if want_file and _norm(want_file) != _norm(file_name):
        return False
    if want_sheet and _norm(want_sheet) != _norm(sheet_name):
        return False
    return True


def _value_evidence(
    files: dict[str, Any], query: str, scope: Optional[dict]
) -> dict[tuple[str, str], list[dict[str, str]]]:
    """Sampled values that appear literally in the query, per sheet.

    This is the value-level index doing its job: a sheet that actually
    contains 'BESTEX' in a sampled column outranks one that merely has a
    column named Client.
    """
    q = " ".join(query.split()).casefold()
    evidence: dict[tuple[str, str], list[dict[str, str]]] = {}
    for file_name, file_info in files.items():
        for sheet_name, sheet_info in file_info.get("sheets", {}).items():
            if not _scope_allows(scope, file_name, sheet_name):
                continue
            hits = []
            for column, values in sheet_info.get("sampled_values", {}).items():
                for value in values:
                    if len(value) >= 3 and value.casefold() in q:
                        hits.append({"column": column, "value": value})
            if hits:
                hits.sort(key=lambda h: -len(h["value"]))
                evidence[(file_name, sheet_name)] = hits
    return evidence


def _pick_return_column(
    query_tokens: set[str], columns: list[str], exclude: set[str]
) -> Optional[str]:
    """Best lexical match between the query and a column name, if any."""
    best, best_score = None, 0.0
    for column in columns:
        if column in exclude:
            continue
        col_tokens = _tokens(column)
        if not col_tokens:
            continue
        score = len(col_tokens & query_tokens) / len(col_tokens)
        if score > best_score:
            best, best_score = column, score
    return best if best_score >= _RETURN_COLUMN_MIN_OVERLAP else None


def _query_ngrams(query: str, max_words: int = 4) -> list[str]:
    words = [w for w in re.split(r"[^\w]+", query) if w]
    grams = []
    for size in range(1, max_words + 1):
        for i in range(len(words) - size + 1):
            grams.append(" ".join(words[i : i + size]))
    return grams


def _fuzzy_sampled_suggestions(
    files: dict[str, Any], query: str, scope: Optional[dict]
) -> list[str]:
    """Sampled values close to some phrase of the query — the misspelling net."""
    distinct: dict[str, None] = {}
    for file_name, file_info in files.items():
        for sheet_name, sheet_info in file_info.get("sheets", {}).items():
            if not _scope_allows(scope, file_name, sheet_name):
                continue
            for values in sheet_info.get("sampled_values", {}).values():
                for value in values:
                    distinct[value] = None

    grams = _query_ngrams(query)
    scored: list[tuple[float, str]] = []
    for value in distinct:
        value_words = len(value.split())
        best = 0.0
        v = value.casefold()
        for gram in grams:
            if abs(len(gram.split()) - value_words) > 1:
                continue
            matcher = difflib.SequenceMatcher(None, v, gram.casefold())
            if matcher.real_quick_ratio() < _FUZZY_SUGGESTION_CUTOFF:
                continue
            best = max(best, matcher.ratio())
        if best >= _FUZZY_SUGGESTION_CUTOFF and best < 1.0:
            scored.append((best, value))
    scored.sort(key=lambda pair: -pair[0])
    return [value for _, value in scored[:5]]


async def _probe_sheet(
    file_name: str,
    file_info: dict[str, Any],
    sheet_name: str,
    sheet_info: dict[str, Any],
    key_candidates: list[dict[str, str]],
    return_column: str,
) -> dict[str, Any]:
    """Runs the narrow read against one sheet: key column, then matched rows.

    Returns a per-sheet outcome; exceptions propagate to the caller's gather.
    """
    columns = sheet_info.get("columns", [])
    address = sheet_info.get("used_range_address")
    if not address:
        raise ValueError(
            f"'{sheet_name}' in '{file_name}' has no cached range address — "
            f"the graph predates v0.2. Run scan_workspace once to upgrade it."
        )
    col1, row1, col2, row2 = parse_range(address)
    header_row = sheet_info.get("header_row", 1)
    data_start = row1 + header_row  # first data row, 1-based absolute
    item_id = file_info["item_id"]

    outcome: dict[str, Any] = {
        "file": file_name,
        "sheet": sheet_name,
        "found": False,
        "return_column": return_column,
        "read_values": [],
    }
    if data_start > row2 or return_column not in columns:
        return outcome

    best: Optional[tuple[dict[str, str], list[int]]] = None
    distinct_read: dict[str, None] = {}
    for candidate in key_candidates[:_MAX_KEY_PROBES_PER_SHEET]:
        if candidate["column"] not in columns:
            continue
        key_abs_col = col1 + columns.index(candidate["column"])
        key_address = build_range(key_abs_col, data_start, key_abs_col, row2)
        data = await get_range(item_id, sheet_name, key_address, select="values")
        cells = [row[0] if row else None for row in data.get("values", [])]
        for cell in cells:
            if isinstance(cell, str) and cell.strip():
                distinct_read[cell.strip()] = None
        matches = [
            i for i, cell in enumerate(cells) if _key_match(cell, candidate["value"])
        ]
        if matches and (best is None or len(matches) < len(best[1])):
            best = (candidate, matches)
        if best and len(best[1]) == 1:
            break

    outcome["read_values"] = list(distinct_read)
    if best is None:
        return outcome

    key, match_indices = best
    return_idx = columns.index(return_column)
    row_end_col = max(min(col2, col1 + 51), col1 + return_idx)

    matches = []
    for idx in match_indices[:_MAX_AMBIGUOUS_ROWS]:
        abs_row = data_start + idx
        row_data = await get_range(
            item_id,
            sheet_name,
            build_range(col1, abs_row, row_end_col, abs_row),
            select="values,numberFormat",
        )
        row_values = (row_data.get("values") or [[]])[0]
        row_formats = (row_data.get("numberFormat") or [[]])[0]
        converted = []
        for i, cell in enumerate(row_values):
            fmt = row_formats[i] if i < len(row_formats) else None
            converted.append(_serial_to_iso(cell) if is_date_number_format(fmt) else cell)
        matched_row = {
            col: converted[i]
            for i, col in enumerate(columns)
            if i < len(converted) and col != return_column
        }
        matches.append(
            {
                "value": converted[return_idx] if return_idx < len(converted) else None,
                "cell": build_cell(col1 + return_idx, abs_row),
                "matched_row": matched_row,
            }
        )

    outcome.update(
        {
            "found": True,
            "key_column": key["column"],
            "key_value": key["value"],
            "matches": matches,
            "total_matches": len(match_indices),
        }
    )
    return outcome


def _response(
    *,
    value: Any = None,
    found: bool = False,
    confidence: Optional[str] = None,
    provenance: Optional[dict] = None,
    alternatives: Optional[list] = None,
    ambiguity: Optional[str] = None,
    suggestions: Optional[list] = None,
    searched: Optional[list] = None,
    note: Optional[str] = None,
) -> dict[str, Any]:
    response: dict[str, Any] = {
        "value": value,
        "found": found,
        "confidence": confidence,
        "provenance": provenance,
        "alternatives": alternatives or [],
        "ambiguity": ambiguity,
    }
    if suggestions is not None:
        response["suggestions"] = suggestions
    if searched is not None:
        response["searched"] = searched
    if note is not None:
        response["note"] = note
    return response


def _aggregate_outcomes(
    outcomes: list[dict[str, Any]],
    errors: list[dict[str, str]],
    key_value_hint: Optional[str],
    files: dict[str, Any],
    query: Optional[str],
    scope: Optional[dict],
) -> dict[str, Any]:
    searched = [{"file": o["file"], "sheet": o["sheet"]} for o in outcomes] + [
        {"file": e["file"], "sheet": e["sheet"], "error": e["error"]} for e in errors
    ]
    found = [o for o in outcomes if o.get("found")]

    if not found:
        suggestions: dict[str, None] = {}
        if key_value_hint:
            read_values: dict[str, None] = {}
            for o in outcomes:
                for v in o.get("read_values", []):
                    read_values[v] = None
            for close in difflib.get_close_matches(
                key_value_hint, list(read_values), n=5, cutoff=0.6
            ):
                suggestions[close] = None
        if query:
            for s in _fuzzy_sampled_suggestions(files, query, scope):
                suggestions[s] = None
        return _response(
            found=False,
            ambiguity="not_found",
            suggestions=list(suggestions),
            searched=searched,
            note=(
                "The key was not found in any candidate sheet. If a suggestion "
                "looks right, retry with that exact spelling."
            ),
        )

    all_matches = [
        {"value": m["value"], "file": o["file"], "sheet": o["sheet"],
         "cell": m["cell"], "matched_row": m["matched_row"]}
        for o in found
        for m in o["matches"]
    ]

    if any(o["total_matches"] > 1 for o in found):
        return _response(
            found=True,
            confidence="ambiguous",
            alternatives=all_matches,
            searched=searched,
            note=(
                "The key matches more than one row. Never assume the first "
                "match — disambiguate with a more specific key or scope."
            ),
        )

    first = all_matches[0]
    if all(_values_agree(first["value"], m["value"]) for m in all_matches[1:]):
        winner = found[0]
        return _response(
            value=first["value"],
            found=True,
            confidence="high",
            provenance={
                "file": first["file"],
                "sheet": first["sheet"],
                "cell": first["cell"],
                "matched_row": first["matched_row"],
                "key_column": winner.get("key_column"),
                "return_column": winner.get("return_column"),
                "corroborated_by": [
                    {"file": m["file"], "sheet": m["sheet"], "cell": m["cell"],
                     "value": m["value"]}
                    for m in all_matches[1:]
                ],
            },
            searched=searched,
        )

    return _response(
        found=True,
        confidence="conflict",
        alternatives=all_matches,
        searched=searched,
        note=(
            "The key was found in several sheets with DIFFERENT values. "
            "Surface every version to the user — do not pick one."
        ),
    )


async def execute_lookup(
    query: Optional[str] = None,
    key_column: Optional[str] = None,
    key_value: Optional[str] = None,
    return_column: Optional[str] = None,
    folder_path: str = "",
    scope: Optional[dict] = None,
) -> dict[str, Any]:
    folder_path = folder_path.rstrip("/")
    files = _workspace_files(folder_path)

    explicit = all([key_column, key_value, return_column])
    if not explicit and any([key_column, key_value, return_column]):
        raise ValueError(
            "Give either a natural-language query, or all three of "
            "key_column, key_value and return_column."
        )
    if not explicit and not query:
        raise ValueError("Give a query or an explicit key/value/return triple.")

    # ---- pick candidate sheets and per-sheet key candidates, graph-only.
    plan: list[tuple[str, str, list[dict[str, str]], str]] = []
    key_value_hint = key_value

    if explicit:
        with_columns = [
            (fn, sn, si)
            for fn, fi in files.items()
            for sn, si in fi.get("sheets", {}).items()
            if _scope_allows(scope, fn, sn)
            and key_column in si.get("columns", [])
            and return_column in si.get("columns", [])
        ]
        if not with_columns:
            all_columns = sorted(
                {c for fi in files.values() for si in fi.get("sheets", {}).values()
                 for c in si.get("columns", [])}
            )
            raise ValueError(
                f"Column(s) '{key_column}' and '{return_column}' not found "
                f"together in any sheet. Columns present in this workspace "
                f"include: {all_columns[:40]}"
            )
        # Sheets whose sampled values contain the key rank first — evidence
        # beats enumeration.
        evidenced = [
            entry
            for entry in with_columns
            if any(
                _norm(v) == _norm(key_value)
                for v in entry[2].get("sampled_values", {}).get(key_column, [])
            )
        ]
        chosen = evidenced or with_columns
        truncated = len(chosen) > _MAX_EXPLICIT_CANDIDATES
        for fn, sn, si in chosen[:_MAX_EXPLICIT_CANDIDATES]:
            plan.append(
                (fn, sn, [{"column": key_column, "value": key_value}], return_column)
            )
    else:
        truncated = False
        query_tokens = _tokens(query)
        evidence = _value_evidence(files, query, scope)
        ranked = sorted(
            evidence.items(),
            key=lambda item: (-len(item[1]), -len(item[1][0]["value"]), item[0]),
        )
        for (fn, sn), hits in ranked:
            sheet_info = files[fn]["sheets"][sn]
            ret = _pick_return_column(
                query_tokens,
                sheet_info.get("columns", []),
                exclude={h["column"] for h in hits},
            )
            if not ret:
                continue
            plan.append((fn, sn, hits, ret))
            if len(plan) >= _MAX_CANDIDATE_SHEETS:
                break
        if plan:
            key_value_hint = plan[0][2][0]["value"]

        if not plan:
            suggestions = _fuzzy_sampled_suggestions(files, query, scope)
            if suggestions:
                return _response(
                    found=False,
                    ambiguity="not_found",
                    suggestions=suggestions,
                    note=(
                        "No sampled value in the workspace appears literally in "
                        "the query, but some are close — retry with the exact "
                        "spelling, or pass key_column/key_value/return_column "
                        "explicitly."
                    ),
                )
            candidates = [
                {"file": r.get("file"), "sheet": r.get("sheet"),
                 "score": round(r.get("score", 0.0), 4)}
                for r in search(folder_path, query, n_results=5)
            ]
            return _response(
                found=False,
                ambiguity="routing",
                suggestions=[],
                searched=candidates,
                note=(
                    "The query could not be resolved to a key value in any "
                    "sheet. Pass key_column/key_value/return_column explicitly, "
                    "or include a literal value (a client, a material) in the "
                    "query."
                ),
            )

    # ---- run the narrow reads concurrently (bounded by the Graph client's
    # shared semaphore).
    async def run_one(fn: str, sn: str, keys, ret):
        try:
            return await _probe_sheet(fn, files[fn], sn, files[fn]["sheets"][sn], keys, ret)
        except Exception as exc:  # one broken sheet must not sink the lookup
            log(f"[Lookup] {fn}/{sn} failed: {exc}")
            return {"file": fn, "sheet": sn, "error": str(exc)}

    results = await asyncio.gather(*[run_one(*p) for p in plan])
    outcomes = [r for r in results if "error" not in r]
    errors = [r for r in results if "error" in r]

    response = _aggregate_outcomes(
        outcomes, errors, key_value_hint, files, query, scope
    )
    if truncated:
        response["note"] = (
            (response.get("note") or "")
            + f" Only the first {_MAX_EXPLICIT_CANDIDATES} matching sheets were "
            f"searched; narrow with scope for an exhaustive answer."
        ).strip()
    return response
