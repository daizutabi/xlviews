from __future__ import annotations

from functools import cache
from typing import TYPE_CHECKING

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
