"""The table of the Excel sheet is linked to a Pandas DataFrame.

When there are names in the columns, they are ignored unless the index
is unnamed.
"""

from __future__ import annotations

import re
from itertools import chain, product, takewhile
from typing import TYPE_CHECKING, overload

import numpy as np
import pandas as pd
import xlwings as xw
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

from xlviews import common
from xlviews.axes import set_first_position
from xlviews.decorators import wait_updating
from xlviews.element import Bar, Plot, Scatter
from xlviews.formula import aggregate, const
from xlviews.grid import FacetGrid
from xlviews.range import multirange
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
from xlviews.utils import array_index, iter_columns

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterable, Iterator, Sequence
    from typing import Any, Literal

    from numpy.typing import NDArray
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
        """Create a DataFrame on an Excel sheet.

        Args:
            sheet (Sheet): The sheet object.
            row, column (int): The position of the top-left cell.
            cell (Range): The Range of the top-left cell.
            data (DataFrame, optional): The DataFrame to write to the sheet.
            index (bool): Whether to output the index of the DataFrame.
            index_level (int): The depth of the index when importing data
                from the sheet.
            column_level (int): The depth of the columns when importing data
                from the sheet.
            parent (SheetFrame): The parent SheetFrame.
                The child SheetFrame is placed to the right of the parent SheetFrame.
                The parent SheetFrame represents the information extracted
                from the parent.
            head (SheetFrame): The upper SheetFrame.
                The child SheetFrame is placed below the upper SheetFrame.
                The upper SheetFrame represents the additional information
                of the parent.
            style (bool): Whether to decorate the SheetFrame.
            gray (bool): Whether to decorate the SheetFrame in gray.
            autofit (bool): Whether to autofit the SheetFrame.
            number_format (str): The number format of the SheetFrame.
            font_size (int): The font size of the SheetFrame.
        """
        self.name = name
        self.parent = parent
        self.children = []
        self.head = head
        self.tail = None

        self.table = None  # TODO: type

        self.stats = None
        self.dist = None

        self.columns_names = None

        if self.parent:  # Locate the child frame to the right of the parent frame.
            self.cell = self.parent.get_child_cell()
            self.parent.add_child_frame(self)

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

        # If the column is a hierarchical index and the index is
        # a normal index, # display the column name in the index column.
        if index and data.columns.nlevels > 1 and data.index.nlevels == 1:
            self.columns_names = list(data.columns.names)
            self.cell.options(transpose=True).value = self.columns_names
            self.expand("down").columns.autofit()

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

    @property
    def row(self) -> int:
        """Return the row of the top-left cell."""
        self.update_cell()
        return self.cell.row

    @property
    def column(self) -> int:
        """Return the column of the top-left cell."""
        self.update_cell()
        return self.cell.column

    @property
    def columns(self) -> list:
        """Return the column names."""
        if self.columns_level == 1:
            return self.expand("right").options(ndim=1).value or []

        if self.columns_names:
            idx = [tuple(self.columns_names)]
        elif self.has_index:
            start = self.cell.offset(self.columns_level - 1)
            end = start.offset(0, self.index_level - 1)
            idx = self.sheet.range(start, end).value or []
        else:
            idx = []

        cs = []
        for k in range(self.columns_level):
            rng = self.cell.offset(k, self.index_level).expand("right")
            cs.append(rng.options(ndim=1).value)
        cs = [tuple(c) for c in zip(*cs, strict=True)]

        return [*idx, *cs]

    @property
    def value_columns(self) -> list:
        return self.columns[self.index_level :]

    @property
    def index_columns(self) -> list[str | tuple[str, ...] | None]:
        return self.columns[: self.index_level]

    @property
    def wide_columns(self) -> list[str]:
        start = self.cell.offset(-1, self.index_level)
        end = start.offset(0, len(self.columns) - self.index_level - 1)
        cs = self.sheet.range(start, end).value or []
        return [c for c in cs if c]

    def __contains__(self, item: str) -> bool:
        return item in self.columns

    def __iter__(self) -> Iterator[str | tuple[str, ...] | None]:
        return iter(self.columns)

    @overload
    def index(
        self,
        column: str | tuple,
        *,
        relative: bool = False,
    ) -> int | tuple[int, int]: ...
    @overload
    def index(
        self,
        column: list[str | tuple],
        *,
        relative: bool = False,
    ) -> list[int | tuple[int, int]]: ...
    @overload
    def index(
        self,
        column: dict,
        *,
        relative: bool = False,
    ) -> tuple[int, int]: ...
    def index(
        self,
        column: str | tuple | list[str | tuple] | dict,
        *,
        relative: bool = False,
    ) -> int | tuple[int, int] | list[int | tuple[int, int]]:
        """Return the column index (1-indexed).

        If the column is a hierarchical index and the column name is specified,
        return the row index. If relative is True, return the relative position
        from `self.cell`.
        """
        if isinstance(column, list):
            return [self.index(c, relative=relative) for c in column]

        if isinstance(column, dict):
            return self.index_dict(column, relative=relative)

        if self.columns_names and isinstance(column, str):
            return self.index_row(column, relative=relative)

        columns = self.columns
        offset = 1 if relative else self.column

        if column in columns:
            return columns.index(column) + offset

        return self.index_wide(column, relative=relative)

    def index_row(self, column: str, *, relative: bool = False) -> int:
        if not self.columns_names:
            raise NotImplementedError

        row = self.columns_names.index(column)
        return row + 1 if relative else row + self.row

    @overload
    def index_wide(self, column: str, *, relative: bool = False) -> tuple[int, int]: ...
    @overload
    def index_wide(
        self,
        column: tuple[str, str | float],
        *,
        relative: bool = False,
    ) -> int: ...
    def index_wide(
        self,
        column: str | tuple[str, str | float],
        *,
        relative: bool = False,
    ) -> tuple[int, int] | int:
        value_columns = self.value_columns

        cell_start = self.cell.offset(-1, self.index_level)
        cell_end = cell_start.offset(0, len(value_columns) - 1)
        names = self.sheet.range(cell_start, cell_end).options(ndim=1).value or []

        name = column[0] if isinstance(column, tuple) else column
        start = names.index(name)
        end = len(list(takewhile(lambda n: n is None, names[start + 1 :]))) + start

        offset = self.index_level + (1 if relative else self.cell.column)

        if isinstance(column, str):
            return start + offset, end + offset

        values = value_columns[start : end + 1]
        return values.index(column[1]) + start + offset

    def index_dict(self, column: dict, *, relative: bool = False) -> tuple[int, int]:
        if not self.columns_names:
            raise NotImplementedError

        by = list(column.keys())
        key = tuple(column.values())

        index = self.groupby(by)[key]

        if len(index) != 1 or len(index[0]) != 2:
            raise NotImplementedError

        offset = -self.column + 1 if relative else 0
        return index[0][0] + offset, index[0][1] + offset

    def update_cell(self) -> None:  # for bug fix in cell.row, cell.column, cell.expand
        self.cell = self.cell.offset(0, 0)

    def expand(self, mode: str = "table") -> Range:
        self.update_cell()
        return self.cell.expand(mode)

    @property
    def data(self) -> DataFrame:
        """Return the data as a DataFrame."""
        if self.cell.value is None and self.columns_level > 1:
            rng = self.cell.offset(self.columns_level - 1).expand()
            rng = rng.options(DataFrame, index=self.index_level, header=1)
            df = rng.value
            df.columns = MultiIndex.from_tuples(self.value_columns)
            return df

        rng = self.expand()
        rng = rng.options(DataFrame, index=self.index_level, header=self.columns_level)
        df = rng.value

        if not isinstance(df, DataFrame):
            raise NotImplementedError

        if self.columns_names:
            df.index.name = None
            df.columns.names = self.columns_names

        return df

    @property
    def visible_data(self) -> DataFrame:
        self.update_cell()
        start = self.cell.offset(1, 0)
        end = start.offset(len(self) - 1, len(self.columns) - 1)
        range_ = self.sheet.range(start, end)
        data = range_.api.SpecialCells(xw.constants.CellType.xlCellTypeVisible)
        value = [row.Value[0] for row in data.Rows]
        df = DataFrame(value, columns=self.columns)

        if self.has_index and self.index_level:
            df = df.set_index(list(df.columns[: self.index_level]))

        return df

    def range(
        self,
        column: str | tuple | dict | None = None,
        start: int | Literal[False] | None = None,
        end: int | None = None,
    ) -> Range:
        """Return the range of the column.

        If the column is a hierarchical index and the column name is specified,
        return the range of the column.

        Args:
            column (str or tuple or dict, optional): The name of the column.
                If omitted, return the range of the entire SheetFrame.
                If a dict is specified, filter by the hierarchical column.
            start (int, optional):
                - None: first row
                - False: entire row with column row
                - 0: column row
                - -1: entire row data without column row
                - other: specified row
            end (int, optional):
                - None : same as start.
                - other: specified row
        """
        if column is None:
            return self.range_all()

        if column == "index":
            return self.range_index(start, end)

        return self.range_column(column, start, end)

    def range_all(self) -> Range:
        start = self.cell
        row_offset = self.columns_level + len(self) - 1
        column_offset = self.index_level + len(self.value_columns) - 1
        end = start.offset(row_offset, column_offset)
        return self.sheet.range(start, end)

    def range_index(
        self,
        start: int | Literal[False] | None = None,
        end: int | None = None,
    ) -> Range:
        """Return the range of the index."""
        if not self.index_level:
            raise NotImplementedError

        if self.columns_names and start in [-1, 0, False]:
            raise ValueError("index start must be a specific row")

        c = self.index_level - 1

        match start:
            case None:
                cell_start = self.cell.offset(self.columns_level)
                cell_end = cell_start.offset(0, c)

            case False:
                cell_start = self.cell
                cell_end = cell_start.offset(self.columns_level + len(self) - 1, c)

            case 0:
                cell_start = self.cell
                cell_end = cell_start.offset(self.columns_level - 1, c)

            case -1:
                cell_start = self.cell.offset(self.columns_level)
                cell_end = cell_start.offset(len(self) - 1, c)

            case _:
                column = self.cell.column
                cell_start = self.sheet.range(start, column)
                cell_end = self.sheet.range(end or start, column + c)

        return self.sheet.range(cell_start, cell_end)

    def range_column(
        self,
        column: str | tuple | dict,
        start: int | Literal[False] | None = None,
        end: int | None = None,
    ) -> Range:
        if start is False:
            header = self.range_column(column, 0)
            values = self.range_column(column, -1)
            return self.sheet.range(header[0], values[-1])

        index = self.index(column)

        match start:
            case 0:
                start = end = self.row
                if isinstance(column, tuple) and self.columns_level == 1:
                    start -= 1  # wide column
                else:
                    end += self.columns_level - 1
                    if isinstance(index, tuple):
                        start -= 1

            case -1:
                start = self.row + self.columns_level
                end = start + len(self) - 1

            case None:
                start = self.row + self.columns_level

        if isinstance(index, tuple):  # wide column
            column_start, column_end = index
        else:
            column_start = column_end = index

        cell_start = self.sheet.range(start, column_start)

        if end is None:
            if not isinstance(index, tuple):
                return cell_start

            # wide column
            cell_end = cell_start.offset(0, column_end - column_start)
            return self.sheet.range(cell_start, cell_end)

        cell_end = self.sheet.range(end, column_end)
        return self.sheet.range(cell_start, cell_end)

    def __repr__(self) -> str:
        return repr(self.range()).replace("<Range ", "<SheetFrame ")

    def __str__(self) -> str:
        return str(self.range()).replace("<Range ", "<SheetFrame ")

    def rename(self, columns: dict[str, str]) -> None:
        """Rename the columns of the SheetFrame."""
        for old, new in columns.items():
            self.range(old, 0).value = new

    def drop_duplicates(self, column: str | tuple | Iterable[str | tuple]) -> None:
        columns = [column] if isinstance(column, str | tuple) else list(column)

        for column in columns:
            for cell in reversed(self.range(column, -1)[1:]):
                if cell.value == cell.offset(-1).value:
                    cell.value = None

    def get_address(
        self,
        column: str | tuple | dict,
        *,
        formula: bool = False,
        **kwargs,
    ) -> list[str]:
        """Return the address list of the column.

        Args:
            column (str or tuple or dict): The name of the column.
            formula (bool, optional): Whether to add '=' to the beginning
                of the address.
            kwargs Keyword arguments for the `Range.get_address` method.

        Returns:
            list[str]: The address list of the column.
        """
        addresses = []
        for cell in self.range(column, -1):
            addresses.append(cell.get_address(**kwargs))

        if formula:
            addresses = ["=" + address for address in addresses]

        return addresses

    def add_column(self, column: str, value: Any | None = None) -> Range:
        column_int = self.column + len(self.columns)
        cell = self.sheet.range(self.row, column_int)
        cell.value = column

        rng = self.range(column, -1)

        if value is not None:
            rng.options(transpose=True).value = value

        return rng

    def add_formula_column(
        self,
        rng: Range,
        formula: str,
        number_format: str | None = None,
        autofit: bool = False,
        lhs: str | None = None,
    ) -> None:
        """Add a formula column.

        Args:
            rng (Range): The range of the column.
            formula (str or tuple): The formula.
            number_format (str, optional): The number format.
            autofit (bool): Whether to autofit the width.
            lhs (str, optional): The name of the column to be assigned.
        """
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
            rng.value = formula
        else:
            rng.options(transpose=True).value = formula

        if number_format:
            set_number_format(rng, number_format)

        if autofit:
            rng = rng.sheet.range(rng[0].offset(-1), rng[-1])
            rng.autofit()

    def add_wide_column(
        self,
        column: str,
        values: Iterable[str | float],
        *,
        number_format: str | None = None,
        autofit: bool = True,
        style: bool = False,
    ) -> Range:
        """Add a wide column.

        Args:
            column (str): The name of the wide column.
            values (iterable): The values to be expanded horizontally.
            number_format (str, optional): The number format.
            autofit (bool): Whether to autofit the width.
            style (bool): Whether to style the column.
        """
        if self.columns_level != 1:
            raise NotImplementedError

        cell = self.cell.offset(0, len(self.columns))
        values_list = list(values)
        cell.value = values_list

        header = cell.offset(-1)
        header.value = column

        set_font(header, bold=True, color="#002255")
        set_alignment(header, horizontal_alignment="left")

        cell = cell.sheet.range(cell, cell.offset(0, len(values_list)))
        if number_format:
            set_number_format(cell, number_format)

        if autofit:
            range_ = self.range(column, 0)
            range_.autofit()

        if style:
            self.set_style()

        return cell[0].offset(1)

    @overload
    def __getitem__(self, column: str | tuple) -> Series: ...
    @overload
    def __getitem__(self, column: Sequence[str | tuple]) -> DataFrame: ...
    def __getitem__(
        self,
        column: str | tuple | Sequence[str | tuple],
    ) -> Series | DataFrame:
        """Return the column data.

        If column is a string, return a Series. If column is a list,
        return a DataFrame. The index is ignored.
        """
        if isinstance(column, list):
            return DataFrame({c: self[c] for c in column})

        if isinstance(column, str | tuple):
            row = self.row + self.columns_level
            name, column_index = column, self.index(column)
            start = self.sheet.range(row, column_index)

            if len(self) == 1:
                array = [start.value]
            else:
                end = start.offset(len(self) - 1, 0)
                rng = self.sheet.range(start, end)
                array = rng.options(np.array).value

            return Series(array, name=name)

        raise NotImplementedError

    def __setitem__(self, column: str, value: str | list | tuple) -> None:
        if self.columns_level > 1 or not isinstance(column, str):
            raise NotImplementedError

        if column not in self:
            self.add_column(column)

        rng = self.range(column, -1)

        starts_eq = isinstance(value, str) and value.startswith("=")
        if starts_eq or isinstance(value, tuple):
            self.add_formula_column(rng, *value, lhs=column)
        else:
            rng.options(transpose=True).value = value

    def select(self, **kwargs) -> NDArray[np.bool_]:
        """Return the selection of the SheetFrame.

        Keyword arguments are column names and values. The conditions are as follows:
           - list : the specified elements are selected.
           - tuple : the range of the value.
           - other : the value is selected if it matches.
        """

        def filter_(
            sel: NDArray[np.bool_],
            array: Series,
            value: str | list | tuple,
        ) -> None:
            if isinstance(value, list):
                sel &= array.isin(value)
            elif isinstance(value, tuple):
                sel &= (array >= value[0]) & (array <= value[1])
            else:
                sel &= array == value

        if self.columns_names is None:
            # vertical selection
            sel = np.ones(len(self), dtype=bool)

            for key, value in kwargs.items():
                filter_(sel, self[key], value)

            return sel

        # horizontal selection
        columns = self.value_columns
        sel = np.ones(len(columns), dtype=bool)
        df = DataFrame(columns, columns=self.columns_names)

        for key, value in kwargs.items():
            filter_(sel, df[key], value)

        return sel

    def groupby(
        self,
        by: str | list[str] | None,
        sel: NDArray[np.bool_] | None = None,
    ) -> dict[Hashable, list[list[int]]]:
        """Group by the specified column and return the group key and row number."""
        if by is None:
            if self.columns_names is None:
                values = [None] * len(self)
            else:
                values = [None] * len(self.columns)

        elif self.columns_names is None:
            if isinstance(by, list) or ":" in by:
                by = list(iter_columns(self, by))
            values = self[by]

        else:
            df = DataFrame(self.value_columns, columns=self.columns_names)
            values = df[by]

        index = array_index(values, sel)

        if self.columns_names is None:
            offset = self.row + self.columns_level  # vertical
        else:
            offset = self.column + self.index_level  # horizontal

        for key, value in index.items():
            index[key] = [[x + offset for x in v] for v in value]

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

    def get_number_format(self, column):
        cell = self.range(column)
        return get_number_format(cell)

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

    def autofit(self) -> None:
        """Autofits the width of the SheetFrame."""
        self.range().columns.autofit()

    def set_adjacent_column_width(self, width: float) -> None:
        """Set the width of the adjacent empty column."""
        column = self.column + len(self.columns)
        self.sheet.range(1, column).column_width = width

    def hide(self, *, hidden: bool = True) -> None:
        """Hide the SheetFrame."""
        start = self.column
        end = start + len(self.columns)
        column = self.sheet.range((1, start), (1, end)).api.EntireColumn
        column.Hidden = hidden

    def unhide(self) -> None:
        """Unhide the SheetFrame."""
        self.hide(hidden=False)

    def add_child_frame(self, child: SheetFrame) -> None:
        """Add a child SheetFrame."""
        self.children.append(child)
        child.parent = self

    def get_child_cell(self) -> Range:
        """Get the cell of the child SheetFrame."""
        offset = len(self.columns) + 1
        offset += sum(len(child.columns) + 1 for child in self.children)
        return self.cell.offset(0, offset)

    def get_adjacent_cell(self, offset: int = 0) -> Range:
        """Get the adjacent cell of the SheetFrame."""
        if self.children:
            return self.get_child_cell()

        return self.cell.offset(0, len(self.columns) + 1).offset(0, offset)

    def set_columns_alignment(self, alignment: str) -> None:
        start = self.cell
        end = start.offset(0, len(self.columns) - 1)
        rng = self.sheet.range(start, end)
        set_alignment(rng, alignment)

    def astable(self, header=True, autofit=False):
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

    def unlist(self) -> None:
        """Unlist the SheetFrame."""
        if self.table:
            self.table.Unlist()
            self.filtered_header(clear=True)

    def filtered_header(self, clear: bool = False) -> None:
        """Write the filtered element above the header.

        Args:
            clear (bool, optional): If True, clear the header.
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

    def set_chart_position(self, pos: str = "right") -> None:
        set_first_position(self, pos=pos)

    def scatter(self, *args, **kwargs):
        return Scatter(*args, data=self, **kwargs)

    def plot(self, *args, **kwargs):
        return Plot(*args, data=self, **kwargs)

    def bar(self, *args, **kwargs):
        return Bar(*args, data=self, **kwargs)

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
            columns = iter_columns(self, columns)

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
