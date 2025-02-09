from __future__ import annotations

import re
from functools import cache
from typing import TYPE_CHECKING

import xlwings

if TYPE_CHECKING:
    from xlwings import Range as RangeImpl
    from xlwings import Sheet

    from xlviews.range.range import Range


@cache
def index_to_column_name(n: int) -> str:
    """Return the Excel column name from an integer.

    Examples:
        >>> index_to_column_name(1)
        'A'
        >>> index_to_column_name(26)
        'Z'
        >>> index_to_column_name(27)
        'AA'
        >>> index_to_column_name(731)
        'ABC'
    """
    name = ""
    while n > 0:
        n -= 1
        name = chr(n % 26 + 65) + name
        n //= 26

    return name


@cache
def column_name_to_index(col: str) -> int:
    """Return the index from an Excel column name.

    Examples:
        >>> column_name_to_index("A")
        1
        >>> column_name_to_index("Z")
        26
        >>> column_name_to_index("AA")
        27
        >>> column_name_to_index("ABC")
        731
    """
    index = 0
    for char in col:
        index = index * 26 + (ord(char) - ord("A") + 1)

    return index


def split_book(address: str, sheet: Sheet | None = None) -> tuple[str, str]:
    """Return a tuple of the book name and the rest from the address."""
    if address.startswith("["):
        book_name, sheet_name = address[1:].split("]", 1)
        return book_name, sheet_name

    sheet = sheet or xlwings.sheets.active
    return sheet.book.name, address


def split_sheet(address: str, sheet: Sheet | None = None) -> tuple[str, str]:
    """Return a tuple of the sheet name and the rest from the address."""
    index = address.find("!")
    if index != -1:
        return address[:index], address[index + 1 :]

    sheet = sheet or xlwings.sheets.active
    return sheet.name, address


ADDRESS_PATTERN = re.compile(r"([A-Z]+)(\d+):([A-Z]+)(\d+)|([A-Z]+)(\d+)")


def get_index(address: str) -> tuple[int, int, int, int]:
    """Return the row and column indexes from the address.

    Returns:
        tuple[int, int, int, int]: Row, column, row_end, column_end.

    """
    address = address.replace("$", "").upper()

    if m := ADDRESS_PATTERN.match(address):
        if m.group(1):
            column, row, column_end, row_end = m.groups()[:4]
        else:
            column, row = m.groups()[4:]
            column_end, row_end = column, row
        return (
            int(row),
            column_name_to_index(column),
            int(row_end),
            column_name_to_index(column_end),
        )

    msg = f"Invalid address format: {address}"
    raise ValueError(msg)


def reference(
    cell: str | tuple[int, int] | Range | RangeImpl,
    sheet: Sheet | None = None,
) -> str:
    """Return a reference to a cell with sheet name for chart."""
    if isinstance(cell, str):
        return cell

    if isinstance(cell, tuple):
        if sheet is None:
            raise ValueError("`sheet` is required when `cell` is a tuple")

        cell = sheet.range(*cell)

    return "=" + cell.get_address(include_sheetname=True)
