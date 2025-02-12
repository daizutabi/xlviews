"""DataFrame on an Excel sheet."""

from __future__ import annotations

import re
from functools import partial
from itertools import chain, takewhile
from typing import TYPE_CHECKING, overload

import xlwings
from pandas import DataFrame, Index, MultiIndex, Series
from xlwings import Range as RangeImpl
from xlwings import Sheet
from xlwings.constants import CellType

from xlviews.chart.axes import set_first_position
from xlviews.decorators import suspend_screen_updates
from xlviews.element import Bar, Plot, Scatter
from xlviews.grid import FacetGrid
from xlviews.range.formula import Func, aggregate
from xlviews.range.range import Range, iter_addresses
from xlviews.range.style import set_alignment

from . import modify
from .groupby import GroupBy
from .style import set_frame_style, set_wide_column_style
from .table import Table

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence
    from typing import Any, Literal, Self

    from .dist_frame import DistFrame
    from .stats_frame import StatsFrame


class SheetFrame:
    """Data frame on an Excel sheet."""

    cell: RangeImpl
    sheet: Sheet
    index_level: int
    columns_level: int
    columns_names: list[str] | None = None
    table: Table | None = None

    @suspend_screen_updates
    def __init__(
        self,
        row: int,
        column: int,
        data: DataFrame | Series | None = None,
        index: bool = True,
        sheet: Sheet | None = None,
    ) -> None:
        """Create a DataFrame on an Excel sheet.

        Args:
            row (int): The row index of the top-left cell.
            column (int): The column index of the top-left cell.
            data (DataFrame, optional): The DataFrame to write to the sheet.
            index (bool, optional): Whether to output the index of the DataFrame.
            sheet (Sheet, optional): The sheet object.
        """
        self.sheet = sheet or xlwings.sheets.active
        self.cell = self.sheet.range(row, column)

        if data is not None:
            self.set_data(data, index=index)

    def set_data(self, data: DataFrame | Series, index: bool = True) -> None:
        """Set the data of the SheetFrame.

        Args:
            data (DataFrame, optional): The DataFrame to write to the sheet.
            index (bool): Whether to output the index of the DataFrame.
        """
        if isinstance(data, Series):
            data = data.to_frame()

        self.index_level = data.index.nlevels if index else 0
        self.columns_level = data.columns.nlevels

        self.cell.options(DataFrame, index=index).value = data

        if index and data.columns.nlevels > 1 and data.index.nlevels == 1:
            self.columns_names = list(data.columns.names)
            self.cell.options(transpose=True).value = self.columns_names

    def load(self, index_level: int = 1, columns_level: int = 1) -> Self:
        """Load the data of the SheetFrame from Excel sheet.

        Args:
            index_level (int): The depth of the index.
            columns_level (int): The depth of the columns.
        """
        self.index_level = index_level
        self.columns_level = columns_level

        if self.columns_level > 1 and index_level == 1:
            start = self.cell
            end = start.offset(self.columns_level - 1)
            self.columns_names = self.sheet.range(start, end).value

        for api in self.sheet.api.ListObjects:
            if api.Range.Row == self.row and api.Range.Column == self.column:
                self.table = Table(api=api, sheet=self.sheet)
                break

        return self

    def _update_cell(self) -> None:  # important
        self.cell = self.cell.offset()

    def expand(self, mode: str = "table") -> RangeImpl:
        self._update_cell()

        if self.cell.value is None and self.columns_level > 1:
            end = self.cell.offset(self.columns_level - 1).expand(mode)
            return self.sheet.range(self.cell, end)

        return self.cell.expand(mode)

    def __repr__(self) -> str:
        return repr(self.expand()).replace("<Range ", "<SheetFrame ")

    def __str__(self) -> str:
        return str(self.expand()).replace("<Range ", "<SheetFrame ")

    def __len__(self) -> int:
        start = self.cell.offset(self.columns_level)
        cell = start

        while cell.value is not None:
            cell = cell.expand("down")[-1].offset(1)

        return cell.row - start.row

    @property
    def row(self) -> int:
        """Return the row of the top-left cell."""
        self._update_cell()
        return self.cell.row

    @property
    def column(self) -> int:
        """Return the column of the top-left cell."""
        self._update_cell()
        return self.cell.column

    @property
    def columns(self) -> list:
        """Return the column names."""
        if self.columns_level == 1:
            return self.expand("right").options(ndim=1).value or []

        if self.columns_names:
            idx = [tuple(self.columns_names)]
        elif self.index_level:
            start = self.row + self.columns_level - 1, self.column
            end = start[0], start[1] + self.index_level - 1
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
        start = self.row - 1, self.column + self.index_level
        end = start[0], start[1] + len(self.columns) - self.index_level - 1
        cs = self.sheet.range(start, end).value or []
        return [c for c in cs if c]

    def __contains__(self, item: str | tuple) -> bool:
        return item in self.columns

    def __iter__(self) -> Iterator[str | tuple[str, ...] | None]:
        return iter(self.columns)

    @property
    def data(self) -> DataFrame:
        """Return the data as a DataFrame."""
        if self.cell.value is None and self.columns_level > 1:
            rng = self.cell.offset(self.columns_level - 1).expand()
            rng = rng.options(DataFrame, index=self.index_level, header=1)
            df = rng.value
            df.columns = MultiIndex.from_tuples(self.value_columns)
            return df

        rng = self.expand().options(
            DataFrame,
            index=self.index_level,
            header=self.columns_level,
        )
        df = rng.value

        if not isinstance(df, DataFrame):
            raise NotImplementedError

        if self.columns_names:
            df.index.name = None
            df.columns.names = self.columns_names

        return df

    @property
    def visible_data(self) -> DataFrame:
        """Return the visible data as a DataFrame."""
        start = self.row + 1, self.column
        end = start[0] + len(self) - 1, start[1] + len(self.columns) - 1

        rng = self.sheet.range(start, end)
        data = rng.api.SpecialCells(CellType.xlCellTypeVisible)
        value = [row.Value[0] for row in data.Rows]
        df = DataFrame(value, columns=self.columns)

        if self.index_level:
            df = df.set_index(list(df.columns[: self.index_level]))

        return df

    @overload
    def index(self, columns: str | tuple) -> int | tuple[int, int]: ...

    @overload
    def index(
        self,
        columns: Sequence[str | tuple],
    ) -> list[int] | list[tuple[int, int]]: ...

    def index(
        self,
        columns: str | tuple | Sequence[str | tuple],
    ) -> int | tuple[int, int] | list[int] | list[tuple[int, int]]:
        """Return the column index (1-indexed)."""
        if isinstance(columns, str | tuple):
            return self.index([columns])[0]

        if self.columns_names:
            columns_str = [c for c in columns if isinstance(c, str)]
            if len(columns_str) == len(columns):
                return self._index_row(columns_str)

        idx = []
        columns_ = self.columns
        offset = self.column

        for column in columns:
            if column in columns_:
                idx.append(columns_.index(column) + offset)
            else:
                idx.append(self._index_wide(column))

        return idx

    def _index_row(self, columns: list[str]) -> list[int]:
        if not self.columns_names:
            raise ValueError("columns names are not specified")

        columns_names = self.columns_names
        offset = self.row
        return [columns_names.index(c) + offset for c in columns]

    def _index_wide(
        self,
        column: str | tuple[str, str | float],
    ) -> tuple[int, int] | int:
        value_columns = self.value_columns

        start = self.row - 1, self.column + self.index_level
        end = start[0], start[1] + len(value_columns) - 1
        names = self.sheet.range(start, end).options(ndim=1).value or []

        name = column[0] if isinstance(column, tuple) else column
        start = names.index(name)
        end = len(list(takewhile(lambda n: n is None, names[start + 1 :]))) + start

        offset = self.index_level + self.cell.column

        if isinstance(column, str):
            return start + offset, end + offset

        values = value_columns[start : end + 1]

        return values.index(column[1]) + start + offset

    @overload
    def column_index(self, columns: str) -> int: ...

    @overload
    def column_index(self, columns: list[str] | None) -> list[int]: ...

    def column_index(self, columns: str | list[str] | None) -> int | list[int]:
        if self.columns_level != 1:
            raise NotImplementedError

        if isinstance(columns, str):
            return self.column_index([columns])[0]

        column = self.column
        if columns is None:
            columns = self.value_columns
            start = column + self.index_level
            end = start + len(columns)
            return list(range(start, end))

        cs = self.columns
        return [cs.index(c) + column for c in columns]

    def range(
        self,
        column: str | tuple,
        offset: Literal[0, -1] | None = None,
    ) -> Range:
        """Return the range of the column.

        Args:
            column (str or tuple): The name of the column.
            offset (int, optional):
                - None: entire row data without column row
                - 0: first row
                - -1: column row
        """
        if self.columns_names and isinstance(column, str):
            raise NotImplementedError

        index = self.index(column)

        match offset:
            case 0:  # first data row
                start = end = self.row + self.columns_level
                if not isinstance(index, tuple):
                    return Range((start, index), sheet=self.sheet)

            case -1:  # column row
                start = end = self.row
                if isinstance(column, tuple) and self.columns_level == 1:
                    start -= 1  # wide column
                else:
                    end += self.columns_level - 1
                    if isinstance(index, tuple):
                        start -= 1

            case None:  # entire data rows
                start = self.row + self.columns_level
                end = start + len(self) - 1

            case _:
                msg = f"invalid offset: {offset}"
                raise ValueError(msg)

        if isinstance(index, tuple):  # wide column
            column_start, column_end = index
        else:
            column_start = column_end = index

        return Range((start, column_start), (end, column_end), sheet=self.sheet)

    @overload
    def column_range(
        self,
        columns: str,
        offset: Literal[0, -1] | None = None,
    ) -> Range: ...

    @overload
    def column_range(
        self,
        columns: list[str] | None,
        offset: Literal[0, -1] | None = None,
    ) -> list[Range]: ...

    def column_range(
        self,
        columns: str | list[str] | None,
        offset: Literal[0, -1] | None = None,
    ) -> Range | list[Range]:
        if self.columns_level != 1:
            raise NotImplementedError

        if isinstance(columns, str):
            return self.column_range([columns], offset)[0]

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

        idx = self.column_index(columns)
        return [Range((start, i), (end, i), sheet=self.sheet) for i in idx]

    def add_column(
        self,
        column: str,
        value: Any | None = None,
        *,
        number_format: str | None = None,
        autofit: bool = False,
        style: bool = False,
    ) -> RangeImpl:
        column_int = self.column + len(self.columns)
        self.sheet.range(self.row, column_int).value = column

        rng = self.range(column).impl

        if value is not None:
            rng.options(transpose=True).value = value
            if number_format:
                rng.number_format = number_format

        if autofit:
            self.sheet.range(rng.offset(-1), rng.last_cell).autofit()

        if style:
            self.style()

        return rng

    def add_formula_column(
        self,
        rng: Range | RangeImpl | str,
        formula: str,
        *,
        number_format: str | None = None,
        autofit: bool = False,
        style: bool = False,
    ) -> RangeImpl:
        """Add a formula column.

        Args:
            rng (Range): The range of the column.
            formula (str or tuple): The formula.
            number_format (str, optional): The number format.
            autofit (bool): Whether to autofit the width.
        """
        columns = self.columns
        wide_columns = self.wide_columns

        if isinstance(rng, str):
            if rng not in columns + wide_columns:
                rng = self.add_column(rng)
            else:
                rng = self.range(rng)

        if isinstance(rng, Range):
            rng = rng.impl

        refs = {}
        for m in re.finditer(r"{(.+?)}", formula):
            column = m.group(1)

            if column in columns:
                ref = self.range(column, 0)
                addr = ref.get_address(row_absolute=False)

            elif column in wide_columns:
                ref = self.range(column, -1)[0].offset(1)
                addr = ref.get_address(column_absolute=False)

            else:
                ref = self.range(column, 0)[0]
                addr = ref.get_address(column_absolute=False, row_absolute=False)

            refs[column] = addr

        rng.value = formula.format(**refs)

        if number_format:
            rng.number_format = number_format

        if autofit:
            self.sheet.range(rng[0].offset(-1), rng[-1]).autofit()

        if style:
            self.style()

        return rng

    def add_wide_column(
        self,
        column: str,
        values: Iterable[str | float],
        *,
        number_format: str | None = None,
        autofit: bool = True,
        style: bool = False,
    ) -> RangeImpl:
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

        rng = self.cell.offset(0, len(self.columns))
        values_list = list(values)
        rng.value = values_list

        header = rng.offset(-1)
        header.value = column

        set_alignment(header, horizontal_alignment="left")

        rng = self.sheet.range(rng, rng.offset(0, len(values_list)))
        if number_format:
            rng.number_format = number_format

        if autofit:
            self.range(column, -1).impl.autofit()

        if style:
            self.style()

        return rng[0].offset(1)

    def __setitem__(self, column: str | tuple, value: Any) -> RangeImpl:
        if column in self:
            rng = self.range(column).impl
        elif isinstance(column, str):
            rng = self.add_column(column)
        else:
            raise NotImplementedError

        if isinstance(value, str) and value.startswith("="):
            return self.add_formula_column(rng, value)

        rng.options(transpose=True).value = value
        return rng

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

        rngs = self.column_range(columns)

        if columns is None:
            columns = self.value_columns

        agg = partial(
            iter_addresses,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            cellwise=True,
            formula=formula,
        )

        values = [list(agg(r)) for r in rngs]
        df = DataFrame(values, index=columns).T

        if self.index_level and self.index_columns[0]:
            index = self._index_frame()
            if len(index.columns) == 1:
                df.index = Index(index[index.columns[0]])
            else:
                df.index = MultiIndex.from_frame(index)

        return df[columns[0]] if is_str else df

    def _index_frame(self) -> DataFrame:
        start = self.cell.offset(self.columns_level - 1)
        end = start.offset(len(self), self.index_level - 1)
        rng = self.sheet.range(start, end)
        return rng.options(DataFrame).value.reset_index()  # type: ignore

    @overload
    def agg(
        self,
        func: Func | dict,
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
        func: Func | dict | Sequence[Func],
        columns: str | list[str] | None = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Series | DataFrame:
        if self.columns_level != 1:
            raise NotImplementedError

        if isinstance(func, dict):
            columns = list(func.keys())
        elif isinstance(columns, str):
            columns = [columns]

        rngs = self.column_range(columns)

        if columns is None:
            columns = self.value_columns

        agg = partial(
            self._agg_column,
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
            return Series([agg(func, r) for r in rngs], index=columns)

        values = [[agg(f, r) for r in rngs] for f in func]
        return DataFrame(values, index=list(func), columns=columns)

    def _agg_column(
        self,
        func: Func,
        rng: Range,
        **kwargs,
    ) -> str:
        if func == "first":
            rng = rng[0]
            func = None

        return aggregate(func, rng, **kwargs)

    def ranges(self) -> Iterator[Range]:
        if self.columns_names is None:
            start = self.column + self.index_level
            end = start + len(self.value_columns) - 1
            offset = self.row + self.columns_level

            for index in range(len(self)):
                yield Range(
                    (index + offset, start),
                    (index + offset, end),
                    sheet=self.sheet,
                )

        else:
            start = self.row + self.columns_level
            end = start + len(self) - 1
            offset = self.column + self.index_level

            for index in range(len(self.value_columns)):
                yield Range(
                    (start, index + offset),
                    (end, index + offset),
                    sheet=self.sheet,
                )

    def melt(
        self,
        func: Func = None,
        value_name: str = "value",
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame:
        """Unpivot a SheetFrame from wide to long format."""
        if self.columns_names is None:
            raise NotImplementedError

        columns = self.value_columns
        df = DataFrame(columns, columns=self.columns_names)

        agg = partial(
            aggregate,
            func,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        df[value_name] = list(map(agg, self.ranges()))
        return df

    def pivot_table(
        self,
        values: str | list[str] | None = None,
        index: str | list[str] | None = None,
        columns: str | list[str] | None = None,
        aggfunc: Func = None,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame:
        if aggfunc is None:
            data = self.get_address(
                [values] if isinstance(values, str) else values,
                row_absolute=row_absolute,
                column_absolute=column_absolute,
                include_sheetname=include_sheetname,
                external=external,
                formula=formula,
            )

        else:
            if index is None:
                by = []
            else:
                by = [index] if isinstance(index, str) else index
            if columns is None:
                if not by:
                    raise ValueError("index and columns cannot be None")
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

    def get_number_format(self, column: str | tuple) -> str:
        return self.range(column, 0).impl.number_format

    def number_format(
        self,
        number_format: str | dict | None = None,
        *,
        autofit: bool = False,
        **columns_format,
    ) -> Self:
        if isinstance(number_format, str):
            start = self.cell.offset(self.columns_level, self.index_level)
            rng = self.sheet.range(start, self.expand().last_cell)
            rng.number_format = number_format
            if autofit:
                rng.autofit()
            return self

        if isinstance(number_format, dict):
            columns_format.update(number_format)

        for column in chain(self.columns, self.wide_columns):
            if not column:
                continue

            for pattern, number_format in columns_format.items():
                column_name = column if isinstance(column, str) else column[0]

                if re.match(pattern, column_name):
                    rng = self.range(column).impl
                    rng.number_format = number_format
                    if autofit:
                        rng.autofit()
                    break

        return self

    def style(self, *, gray: bool = False, **kwargs) -> Self:
        set_frame_style(self, gray=gray, **kwargs)
        set_wide_column_style(self, gray=gray)
        return self

    def autofit(self) -> Self:
        start = self.cell
        end = start.offset(self.columns_level + len(self), len(self.columns) - 1)
        self.sheet.range(start, end).autofit()
        return self

    def alignment(self, alignment: str) -> Self:
        start = self.cell
        end = start.offset(0, len(self.columns) - 1)
        rng = self.sheet.range(start, end)
        set_alignment(rng, alignment)
        return self

    def set_adjacent_column_width(self, width: float) -> None:
        """Set the width of the adjacent empty column."""
        column = self.column + len(self.columns)
        self.sheet.range(1, column).column_width = width

    def get_adjacent_cell(self, offset: int = 0) -> RangeImpl:
        """Get the adjacent cell of the SheetFrame."""
        return self.cell.offset(0, len(self.columns) + offset + 1)

    def move(self, count: int, direction: str = "down", width: int = 0) -> RangeImpl:
        return modify.move(self, count, direction, width)

    def delete(self, direction: str = "up", *, entire: bool = False) -> None:
        return modify.delete(self, direction, entire=entire)

    def as_table(
        self,
        *,
        const_header: bool = True,
        autofit: bool = True,
        style: bool = True,
    ) -> Self:
        if self.table:
            return self

        if self.columns_level != 1:
            raise NotImplementedError

        self.alignment("left")

        end = self.cell.offset(len(self), len(self.columns) - 1)
        rng = self.sheet.range(self.cell, end)

        table = Table(
            rng,
            autofit=autofit,
            const_header=const_header,
            style=style,
            index_level=self.index_level,
        )
        self.table = table

        return self

    def unlist(self) -> Self:
        if self.table:
            self.table.unlist()
            self.table = None

        return self

    def dist_frame(self, *args, **kwargs) -> DistFrame:
        from .dist_frame import DistFrame

        self.set_adjacent_column_width(1)

        self.dist = DistFrame(self, *args, **kwargs)
        return self.dist

    def stats_frame(self, *args, **kwargs) -> StatsFrame:
        from .stats_frame import StatsFrame

        self.stats = StatsFrame(self, *args, **kwargs)
        return self.stats

    def set_chart_position(self, pos: str = "right") -> None:
        set_first_position(self, pos=pos)

    def scatter(self, *args, **kwargs) -> Scatter:
        return Scatter(*args, data=self, **kwargs)

    def plot(self, *args, **kwargs) -> Plot:
        return Plot(*args, data=self, **kwargs)

    def bar(self, *args, **kwargs) -> Bar:
        return Bar(*args, data=self, **kwargs)

    def grid(self, *args, **kwargs) -> FacetGrid:
        return FacetGrid(self, *args, **kwargs)
