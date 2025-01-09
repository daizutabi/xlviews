from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlwings import Range, Sheet


def multirange(
    sheet: Sheet,
    row: int | list[int | tuple[int, int]],
    column: int | list[int | tuple[int, int]],
) -> Range:
    """Create a discontinuous range.

    Either row or column must be an integer.
    If the other is not an integer, it is treated as a list.
    If index is (int, int), it is a simple range.
    Otherwise, each element of index is an int or (int, int), and they are
    concatenated to create a discontinuous range.

    Args:
        sheet (Sheet): The sheet object.
        row (int, tuple, or list): The row number.
        column (int, tuple, or list): The column number.

    Returns:
        Range: The discontinuous range.
    """
    if isinstance(row, int) and isinstance(column, int):
        return sheet.range(row, column)

    if isinstance(row, int) and isinstance(column, list):
        axis = 0
        index = column
    elif isinstance(column, int) and isinstance(row, list):
        axis = 1
        index = row
    else:
        msg = "Either row or column must be an integer."
        raise TypeError(msg)

    def get_range(start_end: int | tuple[int, int]) -> Range:
        if isinstance(start_end, int):
            start = end = start_end
        else:
            start, end = start_end

        if axis == 0:
            return sheet.range((row, start), (row, end))

        return sheet.range((start, column), (end, column))

    union = sheet.book.app.api.Union

    apis = [get_range(i).api for i in index]
    api = apis[0]

    for r in apis[1:]:
        api = union(api, r)

    return sheet.range(api.Address)


def reference(cell: str | tuple[int, int] | Range, sheet: Sheet | None = None) -> str:
    """Return a reference to a cell."""
    if isinstance(cell, str):
        return cell

    if sheet is None:
        if isinstance(cell, tuple):
            raise ValueError("sheet is required when cell is a tuple")

        sheet = cell.sheet

    return "=" + sheet.range(*cell).get_address(include_sheetname=True)
