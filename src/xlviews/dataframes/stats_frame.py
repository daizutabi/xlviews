from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
import pandas as pd
from pandas import DataFrame
from xlwings.constants import Direction

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.formula import AGG_FUNCS, aggregate
from xlviews.range.range import Range
from xlviews.range.range_collection import RangeCollection
from xlviews.range.style import set_font, set_number_format
from xlviews.utils import iter_columns

from .groupby import GroupBy
from .sheet_frame import SheetFrame

if TYPE_CHECKING:
    from collections.abc import Iterator

    from numpy.typing import NDArray


class StatsGroupBy(GroupBy):
    def ranges(self, column: str) -> Iterator[Range | RangeCollection | None]:
        if column in self.by:
            yield from super().first_ranges(column)

        elif column in self.sf.index_columns:
            yield from [None] * len(self.group)

        else:
            yield from super().ranges(column)

    def iter_formulas(
        self,
        column: str,
        funcs: list[str],
    ) -> Iterator[str]:
        for ranges in self.ranges(column):
            for func in funcs:
                yield get_formula(func, ranges)

    def get_index(self, funcs: list[str]) -> list[str]:
        return funcs * len(self.group)

    def get_columns(
        self,
        funcs: list[str],
        func_column_name: str = "func",
    ) -> list[str]:
        columns = self.sf.columns

        if isinstance(funcs, list):
            columns = [func_column_name, *columns]

        return columns

    def get_values(
        self,
        funcs: list[str],
    ) -> NDArray[np.str_]:
        values = [self.get_index(funcs)] if isinstance(funcs, list) else []

        for column in self.sf.columns:
            it = self.iter_formulas(column, funcs)
            values.append(list(it))

        return np.array(values).T

    def get_frame(
        self,
        funcs: list[str],
        func_column_name: str = "func",
    ) -> DataFrame:
        values = self.get_values(funcs)
        columns = self.get_columns(funcs, func_column_name)
        df = DataFrame(values, columns=columns)
        return df.set_index(columns[: -len(self.sf.value_columns)])

    def get_frame2(
        self,
        funcs: list[str],
        func_column_name: str = "func",
    ) -> DataFrame:
        df = self.agg(funcs, self.sf.value_columns, formula=True, as_address=True)
        df = df.stack(level=1, future_stack=True)  # noqa: PD013

        index = df.index.to_frame()
        index = index.rename(columns={index.columns[-1]: func_column_name})

        for c in self.sf.index_columns:
            if c not in self.by:
                index[c] = ""

        df = pd.concat([index, df], axis=1)
        return df.set_index([func_column_name, *self.sf.index_columns])


def get_formula(
    func: str | Range,
    ranges: Range | RangeCollection | None,
) -> str:
    if not ranges:
        return ""

    if isinstance(ranges, Range):
        return "=" + ranges.get_address()

    formula = aggregate(func, ranges)

    return f"={formula}"


class StatsFrame(SheetFrame):
    parent: SheetFrame

    # @turn_off_screen_updating
    def __init__(
        self,
        parent: SheetFrame,
        funcs: str | list[str] | None = None,
        *,
        by: str | list[str] | None = None,
        table: bool = True,
        default: str = "median",
        func_column_name: str = "func",
        succession: bool = False,
        auto_filter: bool = True,
        **kwargs,
    ) -> None:
        """Create a StatsFrame.

        Args:
            parent (SheetFrame): The sheetframe to be aggregated.
            funcs (str, list of str, optional): The aggregation
                functions to be used. The following functions are supported:
                    'count', 'sum', 'min', 'max', 'mean', 'median', 'std', 'soa%'
                None to use the default functions.
            by (str, list of str, optional): The column names to be grouped by.
            table (bool): If True, the frame is displayed in table format.
            default (str, optional): The default function to be displayed.
            func_column_name (str, optional): The name of the function column.
            succession (bool, optional): If True, the continuous index is hidden.
            auto_filter (bool): If True, the displayed functions are limited to
                the default ones.
            **kwargs: Passed to SheetFrame.__init__.
        """
        funcs = get_func(funcs)
        by = get_by(parent, by)
        offset = get_length(parent, by, funcs) + 2

        # Store the position of the parent SheetFrame before moving down.
        row = parent.row
        column = parent.column
        if isinstance(funcs, list):
            column -= 1

        move_down(parent, offset)

        gr = StatsGroupBy(parent, by)
        # df = gr.get_frame(funcs, func_column_name)
        df = gr.get_frame2(funcs, func_column_name)

        super().__init__(
            row,
            column,
            data=df,
            index=parent.has_index,
            autofit=False,
            style=False,
            sheet=parent.sheet,
            **kwargs,
        )
        self.parent = parent

        if table:
            self.as_table(autofit=False, const_header=True)

        self.set_style(autofit=False, succession=succession)

        if isinstance(funcs, list):
            self.set_value_style(func_column_name)

        if table:
            self.set_alignment("left")

        if self.table and auto_filter and isinstance(funcs, list) and len(funcs) > 1:
            func = default if default in funcs else funcs[0]
            self.table.auto_filter(func_column_name, func)

    def set_value_style(self, func_column_name: str) -> None:
        func_index = self.index(func_column_name)

        start = self.column + self.index_level
        end = self.column + len(self.columns)
        columns = [func_index, *range(start, end)]

        get_fmt = self.parent.get_number_format
        formats = [get_fmt(column) for column in self.value_columns]
        formats = [None, *formats]

        group = self.groupby(func_column_name).group

        for key, rows in group.items():
            func = key[0]
            for column, fmt in zip(columns, formats, strict=True):
                rc = RangeCollection(rows, column, self.sheet)

                if func in ["median", "min", "mean", "max", "std", "sum"] and fmt:
                    set_number_format(rc, fmt)

                color = rcParams.get(f"stats.{func}.color")
                italic = rcParams.get(f"stats.{func}.italic")
                set_font(rc, color=color, italic=italic)

                if func == "soa" and column != func_index:
                    set_number_format(rc, "0.0%")

        set_font(self.range(func_column_name), italic=True)


def get_func(func: str | list[str] | None) -> list[str]:
    if func is None:
        func = list(AGG_FUNCS.keys())
        func.remove("sum")
        func.remove("std")
        return func

    return [func] if isinstance(func, str) else func


def get_by(sf: SheetFrame, by: str | list[str] | None) -> list[str]:
    if not by:
        return [c for c in sf.index_columns if isinstance(c, str)]

    return list(iter_columns(sf, by))


def get_length(sf: SheetFrame, by: list[str], funcs: list | dict) -> int:
    n = 1 if isinstance(funcs, dict) else len(funcs)

    if not by:
        return n

    return len(sf.data.reset_index()[by].drop_duplicates()) * n


def has_header(sf: SheetFrame) -> bool:
    start = sf.cell.offset(-1)
    end = start.offset(0, len(sf.columns))
    value = sf.sheet.range(start, end).options(ndim=1).value

    if not isinstance(value, list):
        raise NotImplementedError

    return any(value)


def move_down(sf: SheetFrame, length: int) -> int:
    start = sf.row - 1
    end = sf.row + length - 2

    if has_header(sf):
        end += 1

    rows = sf.sheet.api.Rows(f"{start}:{end}")
    rows.Insert(Shift=Direction.xlDown)
    return end - start + 1
