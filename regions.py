"""Tier 1 of region detection: table bodies derived from a sheet's own formulas.

A real contract workbook does not hold one table per sheet. It holds a banner,
a summary block, four sections with repeated headers, and a totals row — and
the header heuristic in ``structure.detect_header_row`` sees only the first
thing that looks like a header, so everything after it is served to an agent as
data. Summing a quantity column across that gives an answer six times too large,
confidently and without a warning.

The fix that costs nothing: the workbook's own aggregate formulas already say
where each table body starts and ends. ``=SUM(C7:C31)`` in a totals row is the
author stating the extent of the block above it. Extracting those ranges and
merging the adjacent ones recovers every body on the reference workbook, with
the header row landing exactly on ``body_start - 1`` in most cases and one row
above it where the sheet opens with an ``Opening Balance`` line.

Nothing here talks to Graph or to an LLM. It takes a ``formulas`` array and
returns row spans, so it is testable entirely offline.

Row numbers in this module are **absolute 1-based sheet rows**, matching what
the formulas themselves say. That is deliberately *not* the convention used by
``sheet_info["header_row"]``, which is a 1-based offset into the used range;
``structure`` converts between them at the one point where it matters.
"""

import re
from typing import Any, Iterable

# The functions whose range arguments delimit a table body. Measured against
# the reference workbook, these five families locate 11 of 11 bodies across its
# eight sheets. The list is deliberately short: every extra function is another
# chance to pick up a range that spans two sections and silently merge them.
_BODY_FUNCTIONS = (
    "SUM",
    "SUMIF",
    "SUMIFS",
    "COUNT",
    "COUNTA",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGE",
)

_FUNCTION_RE = re.compile(
    r"(?<![A-Za-z0-9_.])(" + "|".join(_BODY_FUNCTIONS) + r")\s*\(",
    re.IGNORECASE,
)

# A1-style range with optional $ anchors: 'C7:C31', '$C$7:$C$31'.
_LOCAL_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([0-9]+):\$?([A-Za-z]{1,3})\$?([0-9]+)$"
)

# A reference carrying a sheet prefix: "'CMC Statement'!C48", "Dashboard!A1:B9",
# "[1]Sheet1!A1" for an external workbook. The bracketed workbook index is
# captured so external references can be discarded rather than mistaken for a
# sibling sheet in this file.
_SHEET_REF_RE = re.compile(
    r"(?:\[(?P<workbook>[^\]]*)\])?"
    r"(?:'(?P<quoted>(?:[^']|'')+)'|(?P<bare>[A-Za-z0-9_.À-￿]+))"
    r"!(?P<address>\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)"
)

# Two blank or sub-total rows inside a block are normal; a wider gap means a
# different section. Anything under four rows is a summary line, not a table.
_MAX_MERGE_GAP = 2
_MIN_BODY_ROWS = 4


def _iter_formulas(formulas: Iterable[Iterable[Any]]) -> Iterable[str]:
    """Yields every formula string in a Graph ``formulas`` array.

    Cells without a formula come back as their literal value (a number, a
    string, an empty string), so anything not starting with '=' is skipped.
    """
    for row in formulas or []:
        for cell in row or []:
            if isinstance(cell, str) and cell.startswith("="):
                yield cell


def _split_arguments(text: str, start: int) -> tuple[list[str], int]:
    """Splits the argument list of a call whose '(' sits at ``start``.

    Returns (arguments, index just past the matching ')'). Commas inside nested
    calls and inside string literals do not split — ``SUMIFS(D7:D31, B7:B31,
    ">=1,000")`` has three arguments, not four. An unbalanced formula (Excel
    permits none, but a truncated read could produce one) yields whatever was
    parsed up to the end of the text.
    """
    args: list[str] = []
    current: list[str] = []
    depth = 0
    in_string = False
    i = start + 1
    while i < len(text):
        ch = text[i]
        if in_string:
            # "" is an escaped quote inside an Excel string literal.
            if ch == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    current.append('""')
                    i += 2
                    continue
                in_string = False
            current.append(ch)
        elif ch == '"':
            in_string = True
            current.append(ch)
        elif ch == "(":
            depth += 1
            current.append(ch)
        elif ch == ")":
            if depth == 0:
                args.append("".join(current))
                return args, i + 1
            depth -= 1
            current.append(ch)
        elif ch == "," and depth == 0:
            args.append("".join(current))
            current = []
        else:
            current.append(ch)
        i += 1
    args.append("".join(current))
    return args, len(text)


def _row_span(argument: str) -> tuple[int, int] | None:
    """The (first_row, last_row) an argument covers, or None if it is not a
    local multi-cell range.

    Rejected on purpose: cross-sheet ranges (a different sheet's geometry says
    nothing about this one), full-column references like ``C:C`` (no rows to
    read), and single cells.
    """
    text = argument.strip()
    if not text or "!" in text or "[" in text:
        return None
    match = _LOCAL_RANGE_RE.match(text)
    if not match:
        return None
    first, last = int(match.group(2)), int(match.group(4))
    if last < first:
        first, last = last, first
    return first, last


def extract_body_spans(formulas: Iterable[Iterable[Any]]) -> list[tuple[int, int]]:
    """Row spans named by the aggregate formulas on one sheet, deduplicated.

    Every range argument of every SUM/COUNT/AVERAGE-family call, in the order
    the rows fall. Overlapping and adjacent spans are left intact here —
    :func:`merge_spans` decides what belongs together.
    """
    spans: set[tuple[int, int]] = set()
    for formula in _iter_formulas(formulas):
        for call in _FUNCTION_RE.finditer(formula):
            args, _ = _split_arguments(formula, call.end() - 1)
            for arg in args:
                span = _row_span(arg)
                if span:
                    spans.add(span)
    return sorted(spans)


def merge_spans(
    spans: Iterable[tuple[int, int]],
    max_gap: int = _MAX_MERGE_GAP,
    min_rows: int = _MIN_BODY_ROWS,
) -> list[tuple[int, int]]:
    """Folds overlapping and near-adjacent spans together, then drops the short ones.

    A single table body is typically named by several formulas — one per
    totalled column, sometimes one per sub-total — covering slightly different
    row ranges. Merging within ``max_gap`` rows reassembles the block. Dropping
    spans under ``min_rows`` discards the two- and three-row summary
    calculations that would otherwise pose as tables.
    """
    ordered = sorted(spans)
    merged: list[list[int]] = []
    for first, last in ordered:
        # first - last - 1 is the count of rows strictly between the two spans:
        # a sub-total line and a blank separator still belong to one block.
        if merged and first - merged[-1][1] - 1 <= max_gap:
            merged[-1][1] = max(merged[-1][1], last)
        else:
            merged.append([first, last])
    return [(a, b) for a, b in merged if b - a + 1 >= min_rows]


def _unescape_sheet_name(name: str) -> str:
    """Excel doubles apostrophes inside a quoted sheet name."""
    return name.replace("''", "'")


def extract_sheet_references(
    formulas: Iterable[Iterable[Any]],
) -> list[dict[str, Any]]:
    """Sheets this sheet's formulas read from, with the addresses they read.

    ``='CMC Statement'!C48`` is the workbook author stating a dependency
    outright — evidence of a completely different quality from the name-token
    and value-overlap agreement that :func:`structure.infer_relationships`
    settles for. References into another *workbook* (``[1]Sheet1!A1``) are
    dropped: this scan only knows about sheets in the file it is reading.

    Returns one entry per referenced sheet, addresses sorted and deduplicated.
    """
    found: dict[str, set[str]] = {}
    for formula in _iter_formulas(formulas):
        for match in _SHEET_REF_RE.finditer(formula):
            if match.group("workbook") is not None:
                continue
            quoted = match.group("quoted")
            name = _unescape_sheet_name(quoted) if quoted else match.group("bare")
            address = match.group("address").replace("$", "")
            found.setdefault(name, set()).add(address)
    return [
        {"sheet": name, "addresses": sorted(addresses)}
        for name, addresses in sorted(found.items())
    ]


def _span_text(first: int, last: int) -> str:
    return f"{first}:{last}"


def build_regions(
    bodies: Iterable[tuple[int, int]],
    first_row: int,
    last_row: int,
) -> list[dict[str, Any]]:
    """Table regions for a sheet, in sheet order, from its merged body spans.

    ``header_row`` is ``body_start - 1``: the row immediately above the block
    the formulas total. On the reference workbook that is exact for eight of
    eleven bodies and one row late for the other three, all of which open with
    an ``Opening Balance`` row sitting between the header and the first numbered
    entry. Tier 1 cannot tell those apart, so every region it produces is marked
    ``"unconfirmed"`` — an agent reading the graph should be able to see which
    layouts have been interpreted and which were only guessed at.

    Bodies are clamped to the used range and dropped if they fall outside it
    entirely, which is what happens when a formula totals a range on a sheet
    that has since been trimmed.
    """
    regions: list[dict[str, Any]] = []
    for body_start, body_end in bodies:
        start = max(body_start, first_row)
        end = min(body_end, last_row)
        if end < start:
            continue
        header_row = start - 1
        if header_row < first_row:
            # The body starts at the top of the used range, so there is no row
            # left to be a header. Report the body without one rather than
            # pointing at a row that does not exist.
            header_row = None
        regions.append(
            {
                "type": "table",
                "range": _span_text(header_row if header_row else start, end),
                "body": _span_text(start, end),
                "header_row": header_row,
                "source": "formula",
                "confidence": "unconfirmed",
            }
        )
    return regions


def unclaimed_rows(
    regions: Iterable[dict[str, Any]],
    first_row: int,
    last_row: int,
) -> list[str]:
    """Row spans no region accounts for, as 'first:last' strings.

    This is the server admitting what it could not classify, and it is exactly
    where the banners, the paired-column summary blocks and the section titles
    live. Tier 2 hands these to an agent to interpret; until then they are the
    honest measure of how much of a sheet is still unexplained.
    """
    claimed: list[tuple[int, int]] = []
    for region in regions:
        try:
            start_text, end_text = str(region.get("range", "")).split(":")
            claimed.append((int(start_text), int(end_text)))
        except (ValueError, AttributeError):
            continue
    claimed.sort()

    gaps: list[str] = []
    cursor = first_row
    for start, end in claimed:
        if start > cursor:
            gaps.append(_span_text(cursor, start - 1))
        cursor = max(cursor, end + 1)
    if cursor <= last_row:
        gaps.append(_span_text(cursor, last_row))
    return gaps


def derive_sheet_regions(
    formulas: Iterable[Iterable[Any]],
    first_row: int,
    last_row: int,
) -> tuple[list[dict[str, Any]], list[str]]:
    """(regions, unclaimed_rows) for one sheet, from its formulas alone."""
    regions = build_regions(
        merge_spans(extract_body_spans(formulas)), first_row, last_row
    )
    return regions, unclaimed_rows(regions, first_row, last_row)
