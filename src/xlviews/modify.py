"""Modify a SheetFrame."""

from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings.constants import Direction

from xlviews.decorators import turn_off_screen_updating
from xlviews.utils import int_to_column_name, iter_columns

if TYPE_CHECKING:
    from xlwings import Range, Sheet

    from xlviews.sheetframe import SheetFrame


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


# def get_sheet_cell(sf: SheetFrame, *args) -> tuple[Sheet, Range]:
#     if len(args) == 0:
#         cell = sf.get_adjacent_cell()
#     elif isinstance(args[0], Sheet):
#         cell = common.get_range(*args[1:], sheet=args[0])
#     else:
#         cell = common.get_range(*args)

#     return cell.sheet, cell


# def get_index_level(sf: SheetFrame, columns: list[str]) -> int:
#     index_columns = sf.index_columns
#     level = next(
#         (i for i, c in enumerate(columns) if c not in index_columns), len(columns)
#     )

#     if columns[-1] in index_columns:
#         index_level += 1

#     return index_level


@turn_off_screen_updating
def copy(
    sf: SheetFrame,
    *args,
    columns: list[str] | None = None,
    n=1,
    header_ref=False,
    sort_index=False,
    sel=None,
    rows=None,
    drop_duplicates=False,
    autofit=True,
    sheet: Sheet | None = None,
    **kwargs,
) -> SheetFrame:
    """
    自分の参照コピーを作成する。

    Parameters
    ----------
    *args :
        SheetFrameの第一引数。コピー先の場所を指定する。
    columns : list of str, optional
        コピーするカラム名
    n : int, optional
        行の展開数
    header_ref : bool, optional
        ヘッダー行を参照するか
    sort_index : bool, optional
        インデックスをソートするか
    sel : dict or list of bool, optional
        コピーする行を選択する辞書かファンシーインデックスで指定
    rows: list of int, optional
        コピーする行を行番号で指定. 0-index
    drop_duplicates : bool, optional
        重複するインデックスを削除するか
    autofit : bool, optional
        オートフィットするか

    Returns
    -------
    SheetFrame
    """
    sheet = sheet or sf.sheet
    cell = sheet.range(*args)
    include_sheetname = sf.sheet != cell.sheet
    columns = sf.columns if columns is None else list(iter_columns(sf, columns))
    index_columns = sf.index_columns

    for index_level, column in enumerate(columns):
        if column not in index_columns:
            break
    if columns[-1] in index_columns:
        index_level += 1

    index_data = sf[columns[:index_level]]

    if isinstance(sel, dict):
        sel = sf.select(**sel)
    if sel is not None:
        index_data = index_data[sel]
    if rows:
        index_data = index_data[index_data.index.isin(rows)]
    if drop_duplicates:
        index_data = index_data.drop_duplicates()
    if sort_index:
        index_data = index_data.sort_values(columns[:index_level])

    header = []
    header_cell = {}
    for column in columns:
        if column not in sf.columns:
            header.append(column)
        else:
            header_ = sf.range(column, 0)
            header_cell[column] = header_
            if header_ref:
                header_ = header_.get_address(include_sheetname=include_sheetname)
                header.append("=" + header_)
            else:
                header.append(column)
    values = [header]
    for row in index_data.index:
        row_values = []
        for column in columns:
            if column in header_cell:
                ref = header_cell[column].offset(int(row) + 1)
                ref = ref.get_address(include_sheetname=include_sheetname)
                value = "=" + ref
            else:
                value = None
            row_values.append(value)
        row_values = [row_values] * n
        values.extend(row_values)

    cell.value = values

    self_columns = sf.columns
    number_format = {
        column: sf.get_number_format(column)
        for column in columns
        if column in self_columns
    }

    sf = sf.__class__(cell, index_level=index_level, autofit=False, **kwargs)

    sf.set_number_format(autofit=False, **number_format)

    if autofit:
        sf.range().columns.autofit()

    return sf


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
