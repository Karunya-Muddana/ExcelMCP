"""A1-notation helpers: parsing and building Excel range addresses.

Graph returns used-range addresses like ``Sheet1!B3:H500`` (or
``'My Sheet'!A1`` for a single cell). Bounded reads — a ten-row header
window, one key column, one result cell — are built from these, so the
arithmetic is 1-based and inclusive throughout, matching Excel's own.
"""

import re

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def col_to_index(letters: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27."""
    index = 0
    for ch in letters.upper():
        if not "A" <= ch <= "Z":
            raise ValueError(f"Invalid column letters: {letters!r}")
        index = index * 26 + (ord(ch) - ord("A") + 1)
    if index == 0:
        raise ValueError("Empty column letters")
    return index


def index_to_col(index: int) -> str:
    """1 -> 'A', 26 -> 'Z', 27 -> 'AA'."""
    if index < 1:
        raise ValueError(f"Column index must be >= 1, got {index}")
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def parse_range(address: str) -> tuple[int, int, int, int]:
    """'Sheet1!B3:H500' or 'B3' -> (col1, row1, col2, row2), 1-based inclusive.

    A leading sheet prefix (quoted or not) is ignored; a single-cell address
    yields a 1x1 range.
    """
    ref = address.split("!")[-1].strip()
    parts = ref.split(":")
    if not ref or len(parts) > 2:
        raise ValueError(f"Cannot parse range address: {address!r}")
    start = _CELL_RE.match(parts[0].strip())
    end = _CELL_RE.match(parts[-1].strip())
    if not start or not end:
        raise ValueError(f"Cannot parse range address: {address!r}")
    col1, row1 = col_to_index(start.group(1)), int(start.group(2))
    col2, row2 = col_to_index(end.group(1)), int(end.group(2))
    if col2 < col1 or row2 < row1:
        raise ValueError(f"Range address is inverted: {address!r}")
    return col1, row1, col2, row2


def build_range(col1: int, row1: int, col2: int, row2: int) -> str:
    """(2, 3, 8, 500) -> 'B3:H500'."""
    return f"{index_to_col(col1)}{row1}:{index_to_col(col2)}{row2}"


def build_cell(col: int, row: int) -> str:
    """(8, 347) -> 'H347'."""
    return f"{index_to_col(col)}{row}"


def parse_span(text: str) -> tuple[int, int]:
    """'7:28' -> (7, 28). Whitespace around the colon and numbers is tolerated."""
    parts = str(text).split(":")
    if len(parts) != 2:
        raise ValueError(f"Cannot parse row span: {text!r}. Expected 'first:last'.")
    try:
        first, last = int(parts[0].strip()), int(parts[1].strip())
    except ValueError:
        raise ValueError(f"Cannot parse row span: {text!r}. Expected 'first:last'.")
    return first, last


def region_to_df_slice(
    used_range_address: str, header_row: int, body: str
) -> tuple[int, int]:
    """Converts a region's absolute body span into a DataFrame row slice.

    ``used_range_address`` is the sheet's used-range address (e.g.
    ``"Sheet1!C5:H60"``); ``header_row`` is the 1-based offset of the header
    into that used range — the convention ``sheet_info["header_row"]`` already
    uses. ``body`` is a region's ``"body"`` span in absolute 1-based sheet
    rows, e.g. ``"7:28"``.

    A DataFrame built by ``_fetch_sheet_data`` has row 0 sitting on the sheet
    row immediately below the header, so absolute row ``r`` lands at index
    ``r - header_abs - 1`` where ``header_abs = row1 + header_row - 1``.

    Returns ``(start, stop)`` — a 0-based, half-open slice suitable for
    ``df.iloc[start:stop]``. Neither bound is clamped to the frame's actual
    length; callers must clamp against ``len(df)`` themselves, since only they
    know it.
    """
    _, row1, _, _ = parse_range(used_range_address)
    header_abs = row1 + header_row - 1
    body_start, body_end = parse_span(body)
    start = body_start - header_abs - 1
    stop = body_end - header_abs
    return start, stop
