"""Modify a SheetFrame."""

from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings.constants import Direction

from xlviews.utils import int_to_column_name

if TYPE_CHECKING:
    from xlwings import Range

    from xlviews.dataframes.sheet import SheetFrame


def _move_down(sf: SheetFrame, count: int) -> Range:
    start = sf.row - 1
    end = start + count - 1

    if sf.cell.offset(-1).formula:
        end += 1

    rows = sf.sheet.api.Rows(f"{start}:{end}")
    rows.Insert(Shift=Direction.xlDown)

    return sf.sheet.range(start + 1, sf.column)


def _move_right(sf: SheetFrame, count: int, width: int) -> Range:
    start = sf.column - 1
    end = start + count - 1

    start_name = int_to_column_name(start)
    end_name = int_to_column_name(end)
    columns_name = f"{start_name}:{end_name}"

    columns = sf.sheet.api.Columns(columns_name)
    columns.Insert(Shift=Direction.xlToRight)

    if width:
        columns = sf.sheet.api.Columns(columns_name)
        columns.ColumnWidth = width

    return sf.sheet.range(sf.row, start + 1)


def move(sf: SheetFrame, count: int, direction: str = "down", width: int = 0) -> Range:
    """Insert empty rows/columns to move the SheetFrame to the right or down.

    Args:
        count (int): The number of empty rows/columns to insert.
        direction (str): 'down' or 'right'
        width (int, optional): The width of the columns to insert.

    Returns:
        Range: Original cell.
    """

    match direction:
        case "down":
            return _move_down(sf, count)

        case "right":
            return _move_right(sf, count, width)

    raise ValueError("direction must be 'down' or 'right'")


def delete(sf: SheetFrame, direction: str = "up", *, entire: bool = False) -> None:
    """Delete the SheetFrame.

    Args:
        direction (str): 'up' or 'left'
        entire (bool): Whether to delete the entire row/column.
    """
    rng = sf.range()
    start = rng[0].offset(-1, -1)
    end = rng[-1].offset(1, 1)

    if sf.wide_columns:
        start = start.offset(-1)

    api = sf.sheet.range(start, end).api

    match direction:
        case "up":
            if entire:
                api.EntireRow.Delete()
            else:
                api.Delete(Shift=Direction.xlUp)

        case "left":
            if entire:
                api.EntireColumn.Delete()
            else:
                api.Delete(Shift=Direction.xlToLeft)

        case _:
            raise ValueError("direction must be 'up' or 'left'")


# @wait_updating
# def product(self, *args, columns=None, **kwargs):
#     """
#     直積シーﾄフレームを生成する。

#     sf.product(a=[1,2,3], b=[4,5,6])とすると、元のシートフレームを
#     9倍に伸ばして、(1,4), (1,5), ..., (3,6)のデータを追加する.

#     Parameters
#     ----------
#     columns: list
#         積をとるカラム名

#     Returns
#     -------
#     SheetFrame
#     """
#     values = []
#     for value in product(*kwargs.values()):
#         values.append(value)
#     df = pd.DataFrame(values, columns=kwargs.keys())
#     if columns is None:
#         columns = self.columns
#     columns += list(df.columns)
#     length = len(self)
#     sf = self.copy(*args, columns=columns, n=len(df))
#     for column in df:
#         sf[column] = list(df[column]) * length
#     sf.set_style(autofit=True)
#     return sf
