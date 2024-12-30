from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range, Sheet

if TYPE_CHECKING:
    from numpy.typing import NDArray
    from xlwings import Sheet


def multirange(
    sheet: Sheet,
    row: int | list[int] | tuple[int, int],
    column: int | list[int] | tuple[int, int],
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

    if isinstance(row, int):
        axis = 0
        index = column  # type: ignore
    elif isinstance(column, int):
        axis = 1
        index = row  # type: ignore
    else:
        msg = "Either row or column must be an integer."
        raise TypeError(msg)

    def _range(start_end):
        if isinstance(start_end, int):
            start = end = start_end
        else:
            start, end = start_end
        if axis == 0:
            return sheet.range((row, start), (row, end))
        return sheet.range((start, column), (end, column))

    if len(index) == 2 and isinstance(index[0], int) and isinstance(index[1], int):
        index = [index]

    ranges = [_range(i).api for i in index]
    union = sheet.book.app.api.Union
    range_ = ranges[0]
    for r in ranges[1:]:
        range_ = union(range_, r)

    return range_


def multirange_indirect(sheet, row, column):
    """
    不連続範囲でもSLOPE関数などが扱えるようにする。
    戻り値はstr
    """
    ranges = multirange(sheet, row, column)
    address = ",".join(['"' + range_.Address + '"' for range_ in ranges])
    return "N(INDIRECT({" + address + "}))"


def reference(sheet, cell):
    """
    Sheetのセルへの参照を返す。
    cellが文字列であればそのまま返す。
    """
    if isinstance(cell, tuple):
        # TODO: tupleのときどの要素使う？連結する？
        cell = cell[0]
    if not isinstance(cell, str):
        cell = sheet.range(*cell).get_address(include_sheetname=True)
        cell = "=" + cell
    return cell
