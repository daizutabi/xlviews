"""The table of the Excel sheet is linked to a Pandas DataFrame.

When there are names in the columns, they are ignored unless the index
is unnamed.
"""

from __future__ import annotations

import re
from itertools import chain, product
from typing import TYPE_CHECKING

import numpy as np
import pandas as pd
import xlwings as xw
from pandas import DataFrame
from xlwings import Sheet

from xlviews import common
from xlviews.axes import set_first_position
from xlviews.decorators import wait_updating
from xlviews.element import Bar, Plot, Scatter
from xlviews.formula import aggregate, const
from xlviews.grid import FacetGrid
from xlviews.style import (
    get_number_format,
    set_alignment,
    set_border,
    set_fill,
    set_font,
    set_frame_style,
    set_number_format,
    set_table_style,
)
from xlviews.utils import add_validation, array_index, columns_list, multirange

if TYPE_CHECKING:
    from xlwings import Range

    from xlviews.dist import DistFrame
    from xlviews.stats import StatsFrame


class SheetFrame:
    """Data frame on an Excel sheet."""

    sheet: Sheet
    cell: Range
    name: str | None
    has_index: bool
    index_level: int
    columns_level: int
    columns_names: list[str] | None
    parent: SheetFrame | None
    children: list[SheetFrame]
    head: SheetFrame | None
    tail: SheetFrame | None
    stats: StatsFrame | None
    dist: DistFrame | None

    @wait_updating
    def __init__(
        self,
        *args,
        name: str | None = None,
        parent: SheetFrame | None = None,
        head: SheetFrame | None = None,
        data: DataFrame | None = None,
        index: bool = True,
        index_level: int = 1,
        columns_level: int = 1,
        style: bool = True,
        gray: bool = False,
        autofit: bool = True,
        number_format: str | None = None,
        font_size: int | None = None,
        **kwargs,
    ) -> None:
        """
        エクセルのシート上のデータフレームを作成する。

        Parameters
        ----------
        sheet : xlwings.main.Sheet
            シートオブジェクト
        row, column : int
            左上のセルの位置
        cell : xlwings.main.Range
            左上のセルのRange
        data : pandas.DataFrame, optional
            データフレームを指定すると、シートに書き出す。
        index : bool or str
            データフレームのインデックスを出力するか。
        index_level : int
            すでにシート上にあるデータを取り込むときのインデックスの深さ
        column_level : int
            すでにシート上にあるデータを取り込むときのカラムの深さ
        parent : SheetFrame
            親シートフレームを指定する。
            自分自身のシートフレームは親フレームの右横に配置される。
            親フレームから抽出した情報を表す。
        head : SheetFrame
            上位シートフレームを指定する。
            自分自身のシートフレームは上位フレームの下に配置される。
            上位シートフレームの補足情報を付け加えることが目的。
        style : bool
            シートフレームを装飾するか
        gray : bool
            グレー装飾にするか
        """
        self.name = name
        self.parent = parent
        self.children = []
        self.head = head
        self.tail = None

        self.table = None  # TODO: type

        self.stats = None
        self.dist = None

        if self.parent:  # Locate the child frame to the right of the parent frame.
            self.cell = self.parent.get_child_cell()
            self.parent.add_child(self)

        elif self.head:  # Locate the child frame below the head frame.
            row_offset = len(self.head) + self.head.columns_level + 1
            self.cell = self.head.cell.offset(row_offset, 0)
            self.head.tail = self

        else:
            if args and isinstance(args[0], Sheet):
                self.cell = common.get_range(*args[1:], sheet=args[0])
            else:
                self.cell = common.get_range(*args)

        self.sheet = self.cell.sheet

        if data is not None:
            self.set_data(
                data,
                index=index,
                number_format=number_format,
                style=style,
                gray=gray,
                autofit=autofit,
                font_size=font_size,
                **kwargs,
            )
        else:
            self.set_data_from_sheet(
                index=index,
                index_level=index_level,
                columns_level=columns_level,
                number_format=number_format,
            )

    def set_data(
        self,
        data: DataFrame,
        *,
        index: bool = True,
        number_format: str | None = None,
        style: bool = True,
        gray: bool = False,
        autofit: bool = True,
        font_size: int | None = None,
        **kwargs,
    ) -> None:
        self.has_index = index
        self.index_level = len(data.index.names) if index else 0
        self.columns_level = len(data.columns.names)

        self.cell.options(DataFrame, index=index).value = data

        if number_format:
            self.set_number_format(number_format)

        if style:
            self.set_style(gray=gray, autofit=autofit, font_size=font_size, **kwargs)

        if self.head is None:
            self.set_adjacent_column_width(1)

        if self.name:
            book = self.sheet.book
            refers_to = "=" + self.cell.get_address(include_sheetname=True)
            book.names.add(self.name, refers_to)

        # If the column is a hierarchical index and the index is a normal index,
        # display the column name in the index column.
        if data.columns.nlevels > 1 and data.index.nlevels == 1:
            self.columns_names = list(data.columns.names)
            self.cell.options(transpose=True).value = self.columns_names
            self.expand("down").columns.autofit()
        else:
            # Ignore the name of the column index.
            self.columns_names = None

    def set_data_from_sheet(
        self,
        *,
        index: bool = True,
        index_level: int = 1,
        columns_level: int = 1,
        number_format: str | None = None,
    ) -> None:
        if self.name:
            book = self.sheet.book
            self.cell = book.names[self.name].refers_to_range
            self.sheet = self.cell.sheet

        self.has_index = index
        self.index_level = index_level
        self.columns_level = columns_level

        if self.columns_level > 1:
            start = self.cell
            end = start.offset(self.columns_level - 1)
            self.columns_names = self.sheet.range(start, end).value
        else:
            self.columns_names = None

        if number_format:
            self.set_number_format(number_format)

        for table in self.sheet.api.ListObjects:
            if table.Range.Row == self.row and table.Range.Column == self.column:
                self.table = table
                break

    def __len__(self) -> int:
        start = self.cell.offset(self.columns_level)
        cell = start

        while cell.value is not None:
            cell = cell.expand("down")[-1].offset(1)

        return cell.row - start.row

    def __contains__(self, item: str) -> bool:
        return item in self.columns

    def __repr__(self):
        return repr(self.range()).replace("<Range ", "<SheetFrame ")

    def __str__(self):
        return str(self.range()).replace("<Range ", "<SheetFrame ")

    def index(
        self,
        column: str | list[str] | dict | tuple[str, str],
        *,
        relative: bool = False,
    ) -> int | list[int]:
        """Return the column index.

        If the column is a hierarchical index and the column name is specified,
        return the row index. If relative is True, return the relative position
        from `self.cell`.
        """
        columns = self.columns
        offset = 1 if relative else self.column
        if isinstance(column, list):
            # return [columns.index(c) + offset for c in column]
            return [self.index(c, relative=relative) for c in column]
        if isinstance(column, dict):
            return self.index_multicolumn(column, relative=relative)
        if column in columns:
            return columns.index(column) + offset
        if self.columns_level > 1:  # 積層カラムのカラム名指定
            row = columns[0].index(column)
            if relative:
                return row + 1
            return row + self.row
        # wide column
        if relative:
            offset = self.index_level + 1
        else:
            offset = self.cell.column + self.index_level
        if isinstance(column, tuple):
            if len(column) != 2:
                raise ValueError("Wide-columnは長さ2のみ")
            start, end = self.index(column[0], relative=True)
            values = self.sheet.range(
                self.cell.offset(0, start - 1),
                self.cell.offset(0, end - 1),
            ).value
            index = values.index(column[1]) + start
            if relative:
                return index
            return index + self.cell.column - 1
        start = self.cell.offset(-1, self.index_level)
        end = start.offset(0, len(columns) - self.index_level - 1)
        values = self.sheet.range(start, end).value
        if start == end:
            values = [values]
        start = end = index = -1
        for index, value in enumerate(values):
            if value == column:
                start = index
            elif value and start != -1:
                end = index - 1
                break
        if start == -1:
            raise IndexError
        if end == -1:
            end = index
        return [start + offset, end + offset]

    def index_multicolumn(self, column, relative=False):
        # 階層インデックスのフィルタリング
        if self.columns_level == 1:
            raise ValueError("階層カラムのときのみ, 辞書によるインデックスが可能")
        by = list(column.keys())
        key = tuple(column.values())
        column = self.groupby(by)[key]
        if len(column) != 1:
            raise ValueError("連続カラムのみ可能")
        if relative:
            return [c - self.column + 1 for c in column[0]]
        return column[0]

    def __getitem__(self, column):
        """
        カラムデータを返す。
        columnが文字列の場合は、Series、
        columnがリストの場合は、DataFrameを返す。

        インデックスは無視される。
        """
        if column == slice(None, None, None):
            df = self.data
            if self.has_index and self.index_level:
                df.reset_index(inplace=True)
            return df
        if isinstance(column, str) or isinstance(column, tuple):
            row = self.row + self.columns_level
            name, column = column, self.index(column)
            start = self.sheet.range(row, column)
            if len(self) == 1:
                array = [start.value]
            else:
                end = start.offset(len(self) - 1, 0)
                range_ = self.sheet.range(start, end)
                array = range_.options(np.array).value
            return pd.Series(array, name=name)
        values = []
        for c in column:
            values.append(self[c])
        values = list(zip(*values, strict=False))
        df = pd.DataFrame(values, columns=column)
        return df

    def __setitem__(self, column, value):
        """
        列を値を設定する。

        Parameters
        ----------
        column : str
            カラム名
        value : str or list or tuple
            カラムの値
        """
        if self.columns_level > 1:
            raise ValueError("未実装")
        if not isinstance(column, str):
            raise ValueError("未実装")
        try:
            range_ = self.range(column, -1)
        except IndexError:
            column_ = self.column + len(self.columns)
            cell = self.sheet.range(self.row, column_)
            cell.value = column
            range_ = self.range(column, -1)
        if (isinstance(value, str) and value.startswith("=")) or isinstance(
            value,
            tuple,
        ):
            self.add_formula_column(range_, value, column)
        else:
            range_.options(transpose=True).value = value

    def update_cell(self):
        """
        cell.row, cell.column, cell.expandのバグ対策
        """
        self.cell = self.cell.offset(0, 0)

    @property
    def row(self) -> int:
        self.update_cell()
        return self.cell.row

    @property
    def column(self) -> int:
        self.update_cell()
        return self.cell.column

    @property
    def columns(self) -> list[str | tuple[str, ...] | None]:
        """Return the column names. This is equivalent to pd.DataFrame.columns."""
        if self.columns_level == 1:
            values = self.expand("right").value
            if isinstance(values, str):
                return [values]

            return values

        # TODO: when self.columns_names is None

        columns_ = []
        for k in range(self.columns_level):
            cell = self.cell.offset(k)
            columns_.append(cell.expand("right").value)

        return [tuple(column) for column in zip(*columns_, strict=False)]

    @property
    def wide_columns(self):
        start = self.cell.offset(-1, self.index_level)
        end = start.offset(0, len(self.columns) - self.index_level - 1)
        values = self.sheet.range(start, end).value
        if values:
            return [value for value in values if value]
        return []

    @property
    def value_columns(self):
        columns = self.columns
        index_level = self.index_level
        return columns[index_level:]

    @property
    def index_columns(self):
        columns = self.columns
        index_level = self.index_level
        return columns[:index_level]

    @property
    def data(self):
        """データフレームを返す。"""
        df = self.expand().options(pd.DataFrame, index=False).value
        # for pandas-0.21.0
        if self.columns_level == 1 and isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        if self.has_index and self.index_level:
            df.set_index(list(df.columns[: self.index_level]), inplace=True)
        return df

    @property
    def visible_data(self):
        self.update_cell()
        start = self.cell.offset(1, 0)
        end = start.offset(len(self) - 1, len(self.columns) - 1)
        range_ = self.sheet.range(start, end)
        data = range_.api.SpecialCells(xw.constants.CellType.xlCellTypeVisible)
        value = [row.Value[0] for row in data.Rows]
        df = pd.DataFrame(value, columns=self.columns)
        if self.has_index and self.index_level:
            df.set_index(list(df.columns[: self.index_level]), inplace=True)
        return df

    def expand(self, direction=None):
        self.update_cell()
        if direction:
            return self.cell.expand(direction)
        return self.cell.expand()

    def range(self, column=None, start=None, end=None):
        """
        カラムの範囲を返す。
        カラムが階層インデックスで、columnが文字列の場合は、カラム列の範囲を返す。

        Parameters
        ----------
        column : str or tuple or dict, optional
            カラム名。省略するとSheetFrame全体を返す。
            dictのとき、積層カラムでのフィルタリング
        start : int, optional
            -1: カラムデータ全体
            0: カラム行
            None: 最初の行
            all: カラム全体
            それ以外：指定行
        end : int, optional
            None : startと同じ
            それ以外: 指定の行

        Returns
        -------
        xlwingws.range
        """
        if start is all:
            header = self.range(column, 0)
            values = self.range(column, -1)
            return self.sheet.range(header[0], values[-1])
        if column is None:
            start = self.cell
            end = start.offset(
                self.columns_level + len(self) - 1,
                self.index_level + len(self.value_columns) - 1,
            )
            return self.sheet.range(start, end)
        if self.columns_level == 1 or isinstance(column, tuple):
            if start == 0:
                start = self.row
                if isinstance(column, tuple) and self.columns_level == 1:
                    # 階層インデックスでないtupleはwide-column
                    end = self.row - 1
                else:
                    end = self.row + self.columns_level - 1
            elif start is None or start == -1:
                if start == -1:
                    end = self.row + len(self) + self.columns_level - 1
                start = self.row + self.columns_level
            column = self.index(column)
            if isinstance(column, list):  # wide column
                column_start, column_end = column
            else:
                column_start = column_end = column
            start = self.sheet.range(start, column_start)
            if end is None:
                if isinstance(column, list):
                    end = start.offset(0, column_end - column_start)
                    return self.sheet.range(start, end)
                return start
            end = self.sheet.range(end, column_end)
            return self.sheet.range(start, end)
        if self.index_level == 1 and column == "index":
            if start == 0:
                start = self.cell
                end = start.offset(self.columns_level - 1)
            elif start is None:
                return self.cell.offset(self.columns_level)
            elif start == -1:
                start = self.cell.offset(self.columns_level)
                end = start.offset(len(self) - 1)
            return self.sheet.range(start, end)
        # 階層カラムのカラム名指定
        columns = self.columns
        if start == 0:
            start = self.column
            end = self.column + self.index_level - 1
        elif start is None or start == -1:
            if start == -1:
                end = self.column + len(columns) - 1
            start = self.column + self.index_level
        row = self.row + columns[0].index(column)
        start = self.sheet.range(row, start)
        if end is None:
            return start
        end = self.sheet.range(row, end)
        return self.sheet.range(start, end)

    def select(self, **kwargs):
        """
        キーワード引数で指定される条件に応じて各要素が選択されるか否かを
        True, Falseのアレイで返す。

        キーワード引数のキーはカラム名、値は条件。条件は以下のものが指定できる。
           - list : 要素を指定する。
           - tuple : 値の範囲を指定する。
           - 他 : 値の一致
        """

        def filter_(sel_, array_, value_):
            if isinstance(value_, list):
                sel_ &= array_.isin(value_)
            elif isinstance(value_, tuple):
                sel_ &= (array_ >= value_[0]) & (array_ <= value_[1])
            else:
                sel_ &= array_ == value_
            return sel_

        if self.columns_names is None:
            # 縦方向に選択
            sel = np.array([True] * len(self))
            for key, value in kwargs.items():
                array = self[key]
                sel = filter_(sel, array, value)
            return sel
        # 横方向に選択
        sel = np.array([True] * len(self.value_columns))
        df = pd.DataFrame(self.value_columns, columns=self.columns_names)
        for key, value in kwargs.items():
            array = df[key]
            sel = filter_(sel, array, value)
        return sel

    def groupby(self, by, sel=None):
        """
        グルーピングして、キーとその行番号の辞書を返す。
        """
        if by is None:
            if self.columns_names is None:
                values = [None] * len(self)
            else:
                values = [None] * len(self.columns)
        elif self.columns_names is None:
            if isinstance(by, list) or ":" in by:
                by = self.columns_list(by)
            values = self[by]
        else:
            df = pd.DataFrame(self.value_columns, columns=self.columns_names)
            values = df[by]

        index = array_index(values, sel)

        # オフセットの調整
        if self.columns_names is None:
            # 縦方向
            row = self.row + self.columns_level
            for key, value in index.items():
                value = [[y + row for y in x] for x in value]
                index[key] = value
        else:
            # 横方向
            column = self.column + self.index_level
            for key, value in index.items():
                value = [[y + column for y in x] for x in value]
                index[key] = value

        return index

    def aggregate(self, func, column: str, by=None, sel=None, **kwargs):
        column = self.index(column)
        if sel is not None:
            sel = self.select(**sel)
        grouped = self.groupby(by, sel)
        dicts = []
        for key, row in grouped.items():
            d = dict(zip(by, key, strict=False))
            range_ = multirange(self.sheet, row, column)
            d["formula"] = aggregate(func, range_, **kwargs)
            dicts.append(d)
        return pd.DataFrame(dicts)

    def set_number_format(
        self,
        *number_format,
        autofit=False,
        split=True,
        **columns_format,
    ):
        if number_format:
            number_format = number_format[0]
            if isinstance(number_format, dict):
                columns_format.update(number_format)
            else:
                cell = xw.Range(
                    self.cell.offset(self.columns_level, self.index_level),
                    self.range()[-1],
                )
                set_number_format(cell, number_format)
                if autofit:
                    cell.autofit()
                return
        if self.columns_level == 1:
            for column in chain(self.columns, self.wide_columns):
                if column in self.columns and isinstance(column, str) and split:
                    column_ = column.split("_")[0]
                else:
                    column_ = column
                if column_ in columns_format:
                    cell = self.range(column, -1)
                    set_number_format(cell, columns_format[column_])
                    if autofit:
                        cell.autofit()
        elif self.index_level == 1:
            for column in list(self.columns[0]) + ["index"]:
                if column in columns_format:
                    cell = self.range(column, -1)
                    set_number_format(cell, columns_format[column])
                    if autofit:
                        cell.autofit()

    def get_number_format(self, column):
        cell = self.range(column)
        return get_number_format(cell)

    def set_style(self, columns_alignment=None, gray=False, **kwargs):
        set_frame_style(
            self.cell,
            self.index_level,
            self.columns_level,
            len(self),
            len(self.value_columns),
            gray=gray,
            **kwargs,
        )
        wide_columns = self.wide_columns
        edge_color = "#aaaaaa" if gray else 0
        for wide_column in wide_columns:
            range_ = self.range(wide_column, 0)
            set_fill(range_, "#eeeeee" if gray else "#f0fff0")
            er = 3 if wide_column == wide_columns[-1] else 2
            edge_width = [1, er - 1, 1, 1] if gray else [2, er, 2, 2]
            set_border(
                range_,
                edge_width=edge_width,
                inside_width=1,
                edge_color=edge_color,
            )
            if gray:
                set_font(range_, color="#aaaaaa")
        for wide_column in wide_columns:
            range_ = self.range(wide_column, 0).offset(-1)
            set_fill(range_, "#eeeeee" if gray else "#e0ffe0")
            el = 3 if wide_column == wide_columns[0] else 2
            edge_width = [el - 1, 2, 2, 1] if gray else [el, 3, 3, 2]
            set_border(
                range_,
                edge_width=edge_width,
                inside_width=None,
                edge_color=edge_color,
            )
            if gray:
                set_font(range_, color="#aaaaaa")
        if columns_alignment:
            self.set_columns_alignment(columns_alignment)

    def autofit(self):
        self.range().columns.autofit()

    def set_adjacent_column_width(self, width):
        """隣接する空列の幅を設定する。"""
        column = self.column + len(self.columns)
        self.sheet.range(1, column).column_width = width

    def hide(self, *, hidden: bool = True) -> None:
        start = self.column
        end = start + len(self.columns)
        column = self.sheet.range((1, start), (1, end)).api.EntireColumn
        column.Hidden = hidden

    def unhide(self) -> None:
        self.hide(hidden=False)

    def add_child(self, child: SheetFrame) -> None:
        self.children.append(child)
        child.parent = self

    def get_child_cell(self) -> Range:
        offset = len(self.columns) + 1
        offset += sum(len(child.columns) + 1 for child in self.children)
        return self.cell.offset(0, offset)

    def get_adjacent_cell(self, offset=0):
        if self.children:
            return self.get_child_cell()
        return self.cell.offset(0, len(self.columns) + 1).offset(0, offset)

    def to_series(self):
        df = self.data
        if len(df.columns) != 1:
            raise ValueError("カラムの数が1ではない。")
        return df[df.columns[0]]

    def set_columns_alignment(self, alignment):
        start = self.cell
        end = start.offset(0, len(self.columns) - 1)
        columns = self.sheet.range(start, end)
        set_alignment(columns, alignment)

    def astable(self, header=True, autofit=False):
        """
        オートフィルタを設定する。
        """
        if self.columns_level != 1:
            return None
        self.set_columns_alignment("left")
        start = self.cell
        end = start.offset(0, len(self.columns) - 1)
        columns = self.sheet.range(start, end)
        table = self.sheet.range(start, end.offset(len(self)))
        xlsrcrange = xw.constants.ListObjectSourceType.xlSrcRange
        table = self.sheet.api.ListObjects.Add(
            xlsrcrange,
            table.api,
            None,
            xw.constants.YesNoGuess.xlYes,
        )
        if autofit:
            columns.api.EntireColumn.AutoFit()
        if header:
            self.filtered_header()
        self.is_table = True
        self.table = table
        set_table_style(table)
        return table

    def unlist(self):
        if self.table:
            self.table.Unlist()
            self.filtered_header(clear=True)

    def filtered_header(self, clear=False):
        """
        列名の上に、フィルターされた要素が一種類の場合にその値を書き出す。

        Parameters
        ----------
        clear : bool, optional
            Trueのときは、消去。
        """
        start = self.cell.offset(self.columns_level)
        end = start.offset(len(self) - 1)
        column = self.sheet.range(start, end)
        start = self.cell.offset(-1)
        end = start.offset(0, self.index_level - 1)
        header = self.sheet.range(start, end)
        if clear:
            header.value = ""
        else:
            header.value = "=" + const(column)
            set_font(header, size=8, italic=True, color="blue")
            set_alignment(header, "center")

    def columns_list(self, columns):
        """
        ':column' or '::column' 形式を通常のカラム名のリストに変換する。
        ':column'はcolumnを含める、'::column'はcolumnの一つ前まで.
        文字列の場合はリストにする。

        Parameters
        ----------
        columns : str or list of str
            カラム名

        Returns
        -------
        columns : list of str
            カラム名のリスト

        """
        return columns_list(self, columns)

    # chart関連 --------------------------------------------------------------
    def scatter(self, *args, **kwargs):
        return Scatter(*args, data=self, **kwargs)

    def plot(self, *args, **kwargs):
        return Plot(*args, data=self, **kwargs)

    def bar(self, *args, **kwargs):
        return Bar(*args, data=self, **kwargs)

    def drop_duplicates(self, column):
        """
        Barプロットように、連続して重複する値を消す。
        """
        columns = [column] if isinstance(column, str) else list(column)
        for column in columns:
            for cell in reversed(self.range(column, -1)[1:]):
                if cell.value == cell.offset(-1).value:
                    cell.value = ""

    def grid(self, *args, **kwargs):
        return FacetGrid(self, *args, **kwargs)

    def distframe(self, *args, **kwargs):
        from xlviews.dist import DistFrame

        self.dist = DistFrame(self, *args, **kwargs)
        return self.dist

    def statsframe(self, *args, **kwargs):
        from xlviews.stats import StatsFrame

        self.stats = StatsFrame(self, *args, **kwargs)
        return self.stats

    # データ追加
    def add_formula_column(self, range_, formula, lhs):
        """
        数式カラムを追加する。

        Parameters
        ----------
        range_ : xw.Range
        formula : str or tuple
        lhs : str
            代入元のカラム名
        """
        if isinstance(formula, tuple):
            if len(formula) == 2:
                formula, format_ = formula
                autofit = False
            else:
                formula, format_, autofit = formula
        else:
            format_ = False
            autofit = False

        self_columns = self.columns
        if isinstance(formula, str) and formula.startswith("="):
            columns = re.findall(r"{.+?}", formula)
            ref_dict = {}
            for column in columns:
                column = column[1:-1]
                if column in self_columns:
                    ref = self.range(column)
                    ref = ref.get_address(row_absolute=False)
                elif column == lhs:  # 自分自身への代入
                    ref = self.range(column, 0)[0]
                    ref = ref.get_address(column_absolute=False)
                else:  # 他に参照される。このときは同型であることが必要。
                    ref = self.range(column)[0]
                    ref = ref.get_address(column_absolute=False, row_absolute=False)
                ref_dict[column] = ref
            formula = formula.format(**ref_dict)
            range_.value = formula
        else:
            range_.options(transpose=True).value = formula

        if format_:
            range_.api.NumberFormatLocal = format_
        if autofit:
            range_ = range_.sheet.range(range_[0].offset(-1), range_[-1])
            range_.autofit()

    def autofilter(self, *args, **field_criteria):
        """
        キーワード引数で指定される条件に応じてフィルタリングする
        キーワード引数のキーはカラム名、値は条件。条件は以下のものが指定できる。
           - list : 要素を指定する。
           - tuple : 値の範囲を指定する。
           - None : 設定されているフィルタをクリアする
           - 他 : 値の一致

        """
        for field, criteria in zip(args[::2], args[1::2], strict=False):
            field_criteria[field] = criteria

        filter_ = self.table.Range.AutoFilter
        operator = xw.constants.AutoFilterOperator
        for field, criteria in field_criteria.items():
            field = self.index(field, relative=True)
            if isinstance(criteria, list):
                criteria = list(map(str, criteria))
                filter_(
                    Field=field,
                    Criteria1=criteria,
                    Operator=operator.xlFilterValues,
                )
            elif isinstance(criteria, tuple):
                filter_(
                    Field=field,
                    Criteria1=f">={criteria[0]}",
                    Operator=operator.xlAnd,
                    Criteria2=f"<={criteria[1]}",
                )
            elif criteria is None:
                filter_(Field=field)
            else:
                filter_(Field=field, Criteria1=f"{criteria}")

    @wait_updating
    def copy(
        self,
        *args,
        columns=None,
        n=1,
        header_ref=False,
        sort_index=False,
        sel=None,
        rows=None,
        drop_duplicates=False,
        autofit=True,
        **kwargs,
    ):
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
        if len(args) == 0:
            sheet = self.sheet
            cell = self.get_adjacent_cell()
        else:
            sheet, cell = get_sheet_cell_row_column(*args)[:2]
        include_sheetname = self.sheet != sheet

        if columns is None:
            columns = self.columns
        else:
            columns = self.columns_list(columns)

        index_columns = self.index_columns
        for index_level, column in enumerate(columns):
            if column not in index_columns:
                break
        if columns[-1] in index_columns:
            index_level += 1

        index_data = self[columns[:index_level]]

        if isinstance(sel, dict):
            sel = self.select(**sel)
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
            if column not in self.columns:
                header.append(column)
            else:
                header_ = self.range(column, 0)
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
        sf = SheetFrame(cell, index_level=index_level, autofit=False, **kwargs)

        self_columns = self.columns
        number_format = {
            column: self.get_number_format(column)
            for column in columns
            if column in self_columns
        }
        sf.set_number_format(**number_format)
        if autofit:
            sf.autofit()
        return sf

    @wait_updating
    def product(self, *args, columns=None, **kwargs):
        """
        直積シーﾄフレームを生成する。

        sf.product(a=[1,2,3], b=[4,5,6])とすると、元のシートフレームを
        9倍に伸ばして、(1,4), (1,5), ..., (3,6)のデータを追加する.

        Parameters
        ----------
        columns: list
            積をとるカラム名

        Returns
        -------
        SheetFrame
        """
        values = []
        for value in product(*kwargs.values()):
            values.append(value)
        df = pd.DataFrame(values, columns=kwargs.keys())
        if columns is None:
            columns = self.columns
        columns += list(df.columns)
        length = len(self)
        sf = self.copy(*args, columns=columns, n=len(df))
        for column in df:
            sf[column] = list(df[column]) * length
        sf.set_style(autofit=True)
        return sf

    def get_address(self, column, formula=False, **kwargs):
        """
        カラムのアドレスリストを返す。

        Parameters
        ----------
        column : str
            カラム名
        formula : bool, optional
            先頭に'='をつけるかどうか
        kwargs :
            get_address関数に渡すキーワード引数

        Returns
        -------
        list
            アドレス文字列のリスト
        """
        range_ = self.range(column, -1)
        addresses = []
        for cell in range_:
            addresses.append(cell.get_address(**kwargs))
        if formula:
            addresses = ["=" + address for address in addresses]
        return addresses

    def set_chart_position(self, pos="right"):
        set_first_position(self, pos=pos)

    def add_wide_column(
        self,
        column: str,
        values,
        number_format=None,
        autofit=True,
        style=False,
    ):
        """
        横方向に展開するカラムを作成する。

        Parameters
        ----------
        column : str
            ワイドカラムの識別名
        values : listable
            横方向に伸びる値のリスト
        number_format : str, optional
            フォーマット
        autofit : bool
            幅を自動調整するか。
        style : bool
            装飾するか

        Returns
        -------
        xw.Range
            セル
        """
        cell = self.cell.offset(0, len(self.columns))
        cell.value = values
        header = cell.offset(-1)
        header.value = column
        set_font(header, bold=True, color="#002255")
        set_alignment(header, horizontal_alignment="left")
        cell = cell.sheet.range(cell, cell.offset(0, len(values)))
        if number_format:
            set_number_format(cell, number_format)
        if autofit:
            range_ = self.range(column, 0)
            range_.autofit()
        if style:
            self.set_style()
        return cell[0].offset(1)

    def add_validation(
        self,
        ref: str,
        column=None,
        name="valid",
        invalid='""',
        valid_value="○",
        invalid_value="×",
        default=True,
    ):
        """
        ○、×でカラムの値を使うかどうかの作業列を追加する。
        フィッティングする値の選別に用いる。

        Parameters
        ----------
        ref : str
            選択するカラム名
        column : str, optional
            選択先のカラム名
            省略すると、ref+'_'
        name : str
            選択を選ぶカラム名
        invalid : str
            選ばなかった時の値
        valid_value, invalid_value : str
            選択文字列
        default : bool
            デフォルトの選択肢
        """
        if column is None:
            column = ref + "_"
        default = valid_value if default else invalid_value
        if name not in self.columns:
            self[name] = default, None, True
            add_validation(
                self.range(name, -1),
                [valid_value, invalid_value],
                default=default,
            )
        number_format = self.get_number_format(ref)
        self[column] = (f'=IF({{{name}}}="○",{{{ref}}},{invalid})', number_format, True)
        self.set_style()

    def move(self, count, direction="down", width=None):
        """
        空の行/列を挿入することで自分自身を右 or 下に移動する。

        Parameters
        ----------
        count : int
            空きを作る行数
        direction : str
            'down' or 'right'
        width : int, optional
            右方向に追加した時のカラム幅

        Returns
        -------
        xw.Range
           元のセル
        """
        if direction == "down":
            start = self.row - 1
            if self.cell.offset(-1).formula:
                end = start + count + 1
            else:
                end = start + count

            rows = self.sheet.api.Rows(f"{start}:{end}")
            rows.Insert(Shift=xw.constants.Direction.xlDown)
            return self.sheet.range(start + 1, self.column)
        if direction == "right":
            start = self.column - 1
            end = start + count
            start_ = self.sheet.range(1, start).get_address().split("$")[1]
            end_ = self.sheet.range(1, end).get_address().split("$")[1]
            columns = self.sheet.api.Columns(f"{start_}:{end_}")
            columns.Insert(Shift=xw.constants.Direction.xlToRight)
            if width:
                columns = self.sheet.api.Columns(f"{start_}:{end_}")
                columns.ColumnWidth = width
            return self.sheet.range(self.row, start + 1)

    def delete(self, direction="up", entire=False):
        """
        自分自身を消去する。
        Parameters
        ----------
        direction : str
            'up' or 'left'
        entire : bool
            行/列全体を消去するか

        Returns
        -------
        """
        range_ = self.range()
        start, end = range_[0], range_[-1]
        start = start.offset(-1, -1)
        if self.wide_columns:
            start = start.offset(-1)
        end = end.offset(1, 1)
        range_ = xw.Range(start, end).api
        if direction == "up":
            if entire:
                range_.EntireRow.Delete()
            else:
                range_.Delete(Shift=xw.constants.Direction.xlUp)
        elif direction == "left":
            if entire:
                range_.EntireColumn.Delete()
            else:
                range_.Delete(Shift=xw.constants.Direction.xlToLeft)
        else:
            raise ValueError('directionは"up" or "left"', direction)

    def set_manual_input(self, column):
        """
        マニュアル入力可能であることを示すスタイルにする

        Parameters
        ----------
        column : str or list of str
            カラム名

        Returns
        -------
        """
        if isinstance(column, list):
            for column_ in column:
                self.set_manual_input(column_)
            return

        range_ = self.range(column, -1)
        set_fill(range_, "yellow")

    def rename(self, columns):
        """
        Parameters
        ----------
        columns : dict
        """
        for column, new_column in columns.items():
            self.range(column, 0).value = new_column
