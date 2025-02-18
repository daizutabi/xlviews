"""DataFrame on an Excel sheet."""

from __future__ import annotations

import re
from functools import partial
from typing import TYPE_CHECKING, overload

import numpy as np
import pandas as pd
import xlwings
from pandas import DataFrame, Series
from xlwings import Range as RangeImpl
from xlwings import Sheet
from xlwings.constants import Direction

from xlviews.core.address import index_to_column_name
from xlviews.core.formula import Func, aggregate
from xlviews.core.index import Index
from xlviews.core.range import Range, iter_addresses
from xlviews.style import set_alignment
from xlviews.utils import suspend_screen_updates

from .groupby import GroupBy
from .style import set_frame_style, set_wide_column_style
from .table import Table

if TYPE_CHECKING:
    from collections.abc import Iterator, Sequence
    from typing import Any, Literal, Self

    from numpy.typing import NDArray


class SheetFrame:
    """Data frame on an Excel sheet."""

    cell: RangeImpl
    sheet: Sheet
    index: pd.Index
    columns: Index
    table: Table | None = None

    @suspend_screen_updates
    def __init__(
        self,
        row: int,
        column: int,
        data: DataFrame,
        sheet: Sheet | None = None,
    ) -> None:
        """Create a DataFrame on an Excel sheet.

        Args:
            row (int): The row index of the top-left cell.
            column (int): The column index of the top-left cell.
            data (DataFrame): The DataFrame to write to the sheet.
            sheet (Sheet, optional): The sheet object.
        """
        self.sheet = sheet or xlwings.sheets.active
        self.cell = self.sheet.range(row, column)

        self.index = data.index
        self.columns = Index(data.columns)

        self.cell.options(DataFrame).value = data

        if data.columns.nlevels > 1 and data.index.nlevels == 1:
            self.cell.options(transpose=True).value = data.columns.names

    def expand(self) -> Range:
        start = self.row, self.column
        end = start[0] + self.height - 1, start[1] + self.width - 1
        return Range(start, end, sheet=self.sheet)

    def __repr__(self) -> str:
        rng = self.expand()
        cls = self.__class__.__name__
        return f"<{cls} {rng.get_address(external=True)}>"

    def __len__(self) -> int:
        return len(self.index)

    def __contains__(self, item: Any) -> bool:
        if not isinstance(item, str | tuple):
            return False

        return item in self.columns

    def __iter__(self) -> Iterator[str | tuple[str, ...] | None]:
        return iter(self.columns)

    @property
    def row(self) -> int:
        """Return the row of the top-left cell."""
        return self.cell.row

    @property
    def column(self) -> int:
        """Return the column of the top-left cell."""
        return self.cell.column

    @property
    def shape(self) -> tuple[int, int]:
        return len(self), len(self.columns)

    @property
    def height(self) -> int:
        return self.columns.nlevels + len(self)

    @property
    def width(self) -> int:
        return self.index.nlevels + len(self.columns)

    @property
    def value(self) -> DataFrame:
        rng = self.expand().impl
        index = self.index.nlevels
        header = self.columns.nlevels
        df = rng.options(DataFrame, index=index, header=header).value

        if not isinstance(df, DataFrame):
            raise NotImplementedError

        return df

    def get_loc(self, column: str) -> int | tuple[int, int]:
        if column in self.index.names:
            return self.index.names.index(column) + self.column

        return self.columns.get_loc(column, self.column + self.index.nlevels)

    @overload
    def get_indexer(self, columns: list[str] | None) -> list[int]: ...

    @overload
    def get_indexer(self, columns: dict[str, Any]) -> NDArray[np.intp]: ...

    def get_indexer(
        self,
        columns: list[str] | dict[str, Any] | None = None,
        **kwargs,
    ) -> list[int] | NDArray[np.intp]:
        if isinstance(columns, dict) or (columns is None and kwargs):
            return self.columns.get_indexer(
                columns,
                self.column + self.index.nlevels,
                **kwargs,
            )

        column = self.column
        if columns is None:
            columns = self.columns.to_list()
            start = column + self.index.nlevels
            end = start + len(columns)
            return list(range(start, end))

        cs = [*self.index.names, *self.columns]
        return [cs.index(c) + column for c in columns]

    @overload
    def get_range(
        self,
        columns: str,
        offset: Literal[0, -1] | None = None,
    ) -> Range: ...

    @overload
    def get_range(
        self,
        columns: list[str] | None,
        offset: Literal[0, -1] | None = None,
    ) -> list[Range]: ...

    def get_range(
        self,
        columns: str | list[str] | None,
        offset: Literal[0, -1] | None = None,
    ) -> Range | list[Range]:
        if self.columns.nlevels != 1:
            raise NotImplementedError

        match offset:
            case 0:
                start = end = self.row + 1
            case -1:
                start = end = self.row
            case None:
                start = self.row + 1
                end = start + len(self) - 1
            case _:
                msg = f"invalid offset: {offset}"
                raise ValueError(msg)

        if isinstance(columns, str):
            loc = self.get_loc(columns)
            if isinstance(loc, int):
                loc = loc, loc
            return Range((start, loc[0]), (end, loc[1]), self.sheet)

        idx = self.get_indexer(columns)
        return [Range((start, i), (end, i), sheet=self.sheet) for i in idx]

    def iter_ranges(self, axis: Literal[0, 1] = 0) -> Iterator[Range]:
        if axis == 0:
            start = self.row + self.columns.nlevels
            end = start + len(self) - 1
            offset = self.column + self.index.nlevels

            for index in range(len(self.columns)):
                yield Range((start, index + offset), (end, index + offset), self.sheet)

        elif axis == 1:
            start = self.column + self.index.nlevels
            end = start + len(self.columns) - 1
            offset = self.row + self.columns.nlevels

            for index in range(len(self)):
                yield Range((index + offset, start), (index + offset, end), self.sheet)

        else:
            raise ValueError("axis must be 0 or 1")

    def add_column(
        self,
        column: str,
        value: Any | None = None,
        *,
        number_format: str | None = None,
        autofit: bool = False,
        style: bool = False,
    ) -> None:
        if self.columns.nlevels != 1:
            raise NotImplementedError

        index = self.column + self.width
        self.sheet.range(self.row, index).value = column
        self.columns.append(column)

        end = self.row + len(self)
        rng = self.sheet.range((self.row + 1, index), (end, index))

        if value is not None:
            rng.options(transpose=True).value = value
            if number_format:
                rng.number_format = number_format

        if autofit:
            rng = self.sheet.range((self.row, index), (end, index))
            rng.autofit()

        if style:
            self.style()

    def add_wide_column(
        self,
        column: str,
        values: Any,
        *,
        number_format: str | None = None,
        autofit: bool = False,
        style: bool = False,
    ) -> None:
        """Add a wide column.

        Args:
            column (str): The name of the wide column.
            values (iterable): The values to be expanded horizontally.
            number_format (str, optional): The number format.
            autofit (bool): Whether to autofit the width.
            style (bool): Whether to style the column.
        """
        if self.columns.nlevels != 1:
            raise NotImplementedError

        index = self.column + self.width
        rng = self.sheet.range(self.row - 1, index)
        rng.value = column
        set_alignment(rng, horizontal_alignment="center")
        self.columns.append(column, values)

        rng = self.sheet.range((self.row, index), (self.row, index + len(values)))
        rng.value = values
        if number_format:
            rng.number_format = number_format

        if autofit:
            rng.autofit()

        if style:
            self.style()

    def add_formula_column(
        self,
        column: str,
        formula: str,
        *,
        number_format: str | None = None,
        autofit: bool = False,
        style: bool = False,
    ) -> None:
        """Add a formula column.

        Args:
            rng (Range): The range of the column.
            formula (str or tuple): The formula.
            number_format (str, optional): The number format.
            autofit (bool): Whether to autofit the width.
        """
        if self.columns.nlevels != 1:
            raise NotImplementedError

        refs = {}
        for m in re.finditer(r"{(.+?)}", formula):
            key = m.group(1)
            loc = self.get_loc(key)

            if isinstance(loc, int):
                ref = Range(self.row + 1, loc, self.sheet)
                addr = ref.get_address(row_absolute=False)

            else:
                ref = Range((self.row, loc[0]), (self.row, loc[0]), self.sheet)
                addr = ref.get_address(column_absolute=False)

            refs[key] = addr

        if isinstance(column, str) and column not in self.columns:
            self.add_column(column)

        rng = self.get_range(column).impl
        rng.value = formula.format(**refs)

        if number_format:
            rng.number_format = number_format

        if autofit:
            self.sheet.range(rng[0].offset(-1), rng[-1]).autofit()

        if style:
            self.style()

    @overload
    def get_address(
        self,
        columns: str,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series: ...

    @overload
    def get_address(
        self,
        columns: list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame: ...

    def get_address(
        self,
        columns: str | list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series | DataFrame:
        if isinstance(columns, str):
            columns = [columns]
            is_str = True
        else:
            is_str = False

        rngs = self.get_range(columns)

        if columns is None:
            columns = self.columns.to_list()

        agg = partial(
            iter_addresses,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            cellwise=True,
            formula=formula,
        )

        values = np.array([list(agg(r)) for r in rngs]).T
        df = DataFrame(values, index=self.index, columns=columns)

        return df[columns[0]] if is_str else df

    @overload
    def agg(
        self,
        func: Func | dict = None,
        columns: str | list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series: ...

    @overload
    def agg(
        self,
        func: Sequence[Func],
        columns: str | list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame: ...

    def agg(
        self,
        func: Func | dict | Sequence[Func] = None,
        columns: str | list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series | DataFrame:
        if self.columns.nlevels != 1:
            if isinstance(func, Func) and columns is None:
                return self.melt(
                    func,
                    row_absolute=row_absolute,
                    column_absolute=column_absolute,
                    include_sheetname=include_sheetname,
                    external=external,
                    formula=formula,
                )

            raise NotImplementedError

        if isinstance(func, dict):
            columns = list(func.keys())
        elif isinstance(columns, str):
            columns = [columns]

        if isinstance(func, Range | RangeImpl):
            if func.sheet.book.name != self.sheet.book.name:
                raise ValueError("Range is from a different book")
            if func.sheet.name != self.sheet.name:
                raise ValueError("Range is from a different sheet")

        rngs = self.get_range(columns)

        if columns is None:
            columns = self.columns.to_list()

        agg = partial(
            self._agg,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        if isinstance(func, dict):
            it = zip(rngs, func.values(), strict=True)
            return Series([agg(f, r) for r, f in it], index=columns)

        if func is None or isinstance(func, str | Range | RangeImpl):
            name = func if isinstance(func, str) else None
            return Series([agg(func, r) for r in rngs], index=columns, name=name)

        values = [[agg(f, r) for r in rngs] for f in func]
        return DataFrame(values, index=list(func), columns=columns)

    def _agg(self, func: Func, rng: Range, **kwargs) -> str:
        if func == "first":
            rng = rng[0]
            func = None

        return aggregate(func, rng, **kwargs)

    def melt(
        self,
        func: Func = None,
        value_name: str = "value",
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series | DataFrame:
        """Unpivot a SheetFrame from wide to long format."""
        agg = partial(
            aggregate,
            func,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        index = self.columns.nlevels == 1
        df = self.columns.to_frame(index=index)
        df[value_name] = list(map(agg, self.iter_ranges(axis=0)))

        if index:
            s = df[value_name]
            if isinstance(func, str | None):
                s.name = func
            return s

        return df

    def pivot_table(
        self,
        values: str | list[str] | None = None,
        index: str | list[str] | None = None,
        columns: str | list[str] | None = None,
        aggfunc: Func | list[Func] = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame:
        if isinstance(aggfunc, list):
            dfs = [
                self.pivot_table(
                    values,
                    index,
                    columns,
                    f,
                    row_absolute,
                    column_absolute,
                    include_sheetname,
                    external,
                    formula,
                )
                for f in aggfunc
            ]

            keys = [f if isinstance(f, str) else None for f in aggfunc]
            return pd.concat(dfs, axis=1, keys=keys)

        if aggfunc is None:
            data = self.get_address(
                [values] if isinstance(values, str) else values,
                row_absolute=row_absolute,
                column_absolute=column_absolute,
                include_sheetname=include_sheetname,
                external=external,
                formula=formula,
            )

        if index is None:
            by = []
        else:
            by = [index] if isinstance(index, str) else index

        if columns is None:
            if not by:
                raise ValueError("No group keys passed!")
        else:
            by = [*by, columns] if isinstance(columns, str) else by + columns

        data = self.groupby(by).agg(
            aggfunc,
            values,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        return data.pivot_table(values, index, columns, aggfunc=lambda x: x)

    def groupby(self, by: str | list[str] | None, *, sort: bool = True) -> GroupBy:
        return GroupBy(self, by, sort=sort)

    def get_number_format(self, column: str) -> str:
        idx = self.get_loc(column)
        if isinstance(idx, tuple):
            idx = idx[0]
        return self.sheet.range(self.row + self.columns.nlevels, idx).number_format

    def number_format(
        self,
        number_format: str | dict | None = None,
        *,
        autofit: bool = False,
        **columns_format,
    ) -> Self:
        if isinstance(number_format, dict):
            columns_format.update(number_format)

        row_start = self.row + self.columns.nlevels
        row_end = row_start + len(self) - 1

        if self.columns.nlevels == 1:
            for column in [*self.index.names, *self.columns]:
                if not column:
                    continue

                for pattern, number_format in columns_format.items():
                    if re.match(pattern, column):
                        rng = self.get_range(column).impl
                        rng.number_format = number_format
                        if autofit:
                            rng.autofit()
                        break

        elif isinstance(number_format, str):
            for i in self.get_indexer(columns_format):
                rng = self.sheet.range((row_start, i), (row_end, i))
                rng.number_format = number_format
                if autofit:
                    rng.autofit()

        else:
            raise NotImplementedError

        return self

    def style(self, **kwargs) -> Self:
        set_frame_style(self, **kwargs)
        set_wide_column_style(self)
        return self

    def autofit(self) -> Self:
        start = self.cell
        end = start.offset(self.height - 1, self.width - 1)
        self.sheet.range(start, end).autofit()
        return self

    def alignment(self, alignment: str) -> Self:
        start = self.cell
        end = start.offset(0, self.width - 1)
        rng = self.sheet.range(start, end)
        set_alignment(rng, alignment)
        return self

    def set_adjacent_column_width(self, width: float) -> None:
        """Set the width of the adjacent empty column."""
        column = self.column + self.width
        self.sheet.range(1, column).column_width = width

    def get_adjacent_cell(self, offset: int = 0) -> RangeImpl:
        """Get the adjacent cell of the SheetFrame."""
        return self.cell.offset(0, self.width + offset + 1)

    def move(self, count: int, direction: str = "down", width: int = 0) -> None:
        return move(self, count, direction, width)

    def as_table(
        self,
        *,
        const_header: bool = True,
        autofit: bool = True,
        style: bool = True,
    ) -> Table:
        if self.table:
            return self.table

        if self.columns.nlevels != 1:
            raise NotImplementedError

        self.alignment("left")

        end = self.cell.offset(len(self), self.width - 1)
        rng = self.sheet.range(self.cell, end)

        table = Table(
            rng,
            autofit=autofit,
            const_header=const_header,
            style=style,
            index_nlevels=self.index.nlevels,
        )
        self.table = table

        return table

    def unlist(self) -> Self:
        if self.table:
            self.table.unlist()
            self.table = None

        return self


def move(sf: SheetFrame, count: int, direction: str = "down", width: int = 0) -> None:
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


def _move_down(sf: SheetFrame, count: int) -> None:
    start = sf.row - 1
    end = start + count - 1

    if sf.cell.offset(-1).formula:
        end += 1

    rows = sf.sheet.api.Rows(f"{start}:{end}")
    rows.Insert(Shift=Direction.xlDown)

    sf.cell = sf.cell.offset()  # update cell


def _move_right(sf: SheetFrame, count: int, width: int) -> None:
    start = sf.column - 1
    end = start + count - 1

    start_name = index_to_column_name(start)
    end_name = index_to_column_name(end)
    columns_name = f"{start_name}:{end_name}"

    columns = sf.sheet.api.Columns(columns_name)
    columns.Insert(Shift=Direction.xlToRight)

    if width:
        columns = sf.sheet.api.Columns(columns_name)
        columns.ColumnWidth = width

    sf.cell = sf.cell.offset()  # update cell
