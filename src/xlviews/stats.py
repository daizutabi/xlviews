from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
from pandas import DataFrame
from xlwings import Range
from xlwings.constants import Direction

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.formula import AGG_FUNCS, aggregate
from xlviews.frame import SheetFrame
from xlviews.group import GroupBy
from xlviews.range import RangeCollection
from xlviews.style import set_font, set_number_format
from xlviews.utils import iter_columns

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
        funcs: list[str] | dict[str, str],
        wrap: str | None = None,
        default: str = "median",
    ) -> Iterator[str]:
        for ranges in self.ranges(column):
            if isinstance(funcs, dict):
                funcs = [funcs.get(column, default)]

            for func in funcs:
                yield get_formula(func, ranges, wrap)

    def get_index(self, funcs: list[str]) -> list[str]:
        return funcs * len(self.group)

    def get_columns(
        self,
        funcs: list[str] | dict[str, str],
        func_column_name: str = "func",
    ) -> list[str]:
        columns = self.sf.columns

        if isinstance(funcs, list):
            columns = [func_column_name, *columns]

        return columns

    def get_values(
        self,
        funcs: list[str] | dict[str, str],
        wrap: str | dict[str, str] | None = None,
        default: str = "median",
    ) -> NDArray[np.str_]:
        values = [self.get_index(funcs)] if isinstance(funcs, list) else []

        for column in self.sf.columns:
            wrap_ = wrap.get(column) if isinstance(wrap, dict) else wrap
            it = self.iter_formulas(column, funcs, wrap_, default)
            values.append(list(it))

        return np.array(values).T

    def get_frame(
        self,
        funcs: list[str] | dict[str, str],
        wrap: str | dict[str, str] | None = None,
        default: str = "median",
        func_column_name: str = "func",
    ) -> DataFrame:
        values = self.get_values(funcs, wrap, default)
        columns = self.get_columns(funcs, func_column_name)
        df = DataFrame(values, columns=columns)
        return df.set_index(columns[: -len(self.sf.value_columns)])


def get_formula(
    func: str | Range,
    ranges: Range | RangeCollection | None,
    wrap: str | None = None,
) -> str:
    if not ranges:
        return ""

    if isinstance(ranges, Range):
        return "=" + ranges.get_address()

    formula = aggregate(func, ranges)

    if wrap:
        formula = wrap.format(formula)

    return f"={formula}"


class StatsFrame(SheetFrame):
    parent: SheetFrame

    @turn_off_screen_updating
    def __init__(
        self,
        parent: SheetFrame,
        funcs: str | list[str] | dict[str, str] | None = None,
        *,
        by: str | list[str] | None = None,
        table: bool = True,
        wrap: str | dict[str, str] | None = None,
        na: str | list[str] | bool = False,
        null: str | list[str] | bool = False,
        default: str = "median",
        func_column_name: str = "func",
        succession: bool = False,
        auto_filter: bool = True,
        **kwargs,
    ) -> None:
        """Create a StatsFrame.

        Args:
            parent (SheetFrame): The sheetframe to be aggregated.
            funcs (str, list of str, dict, optional): The aggregation
                functions to be used. The following functions are supported:
                    'count', 'sum', 'min', 'max', 'mean', 'median', 'std', 'soa%'
                None to use the default functions.
            by (str, list of str, optional): The column names to be grouped by.
            autofilter (bool): If True, the displayed functions are limited to
                the default ones.
            table (bool): If True, the frame is displayed in table format.
            wrap (str or dict, optional): A string to wrap the aggregation
                functions. {} is replaced with the aggregation functions.
            na (bool): If True, self.wrap = 'IFERROR({},NA())' is used.
            null (bool): If True, self.wrap = 'IFERROR({},"")' is used.
            succession (bool, optional): If True, the continuous index is hidden.
            **kwargs: Passed to SheetFrame.__init__.
        """
        funcs = get_func(funcs)

        # Store the position of the parent SheetFrame before moving down.
        row = parent.row
        column = parent.column
        if isinstance(funcs, list):
            column -= 1

        by = list(iter_columns(parent, by)) if by else []
        offset = get_length(parent, by, funcs) + 2

        move_down(parent, offset)

        gr = StatsGroupBy(parent, by)
        wrap = get_wrap(wrap, na=na, null=null)
        df = gr.get_frame(funcs, wrap, default, func_column_name)

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
            func = "median" if "median" in funcs else funcs[0]
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
                rc = RangeCollection.from_index(self.sheet, rows, column)

                if func in ["median", "min", "mean", "max", "std", "sum"]:
                    set_number_format(rc, fmt)

                color = rcParams.get(f"stats.{func}.color")
                italic = rcParams.get(f"stats.{func}.italic")
                set_font(rc, color=color, italic=italic)

                if func == "soa" and column != func_index:
                    set_number_format(rc, "0.0%")

        set_font(self.range(func_column_name), italic=True)


def get_wrap(
    wrap: str | dict[str, str] | None = None,
    *,
    na: str | list[str] | bool = False,
    null: str | list[str] | bool = False,
) -> str | dict[str, str] | None:
    if wrap:
        return wrap

    if na is True:
        return "IFERROR({},NA())"

    if null is True:
        return 'IFERROR({},"")'

    wrap = {}

    if na:
        nas = [na] if isinstance(na, str) else na
        for na in nas:
            wrap[na] = "IFERROR({},NA())"

    if null:
        nulls = [null] if isinstance(null, str) else null
        for null in nulls:
            wrap[null] = 'IFERROR({},"")'

    return wrap or None


def get_func(
    func: str | list[str] | dict[str, str] | None,
) -> list[str] | dict[str, str]:
    if func is None:
        func = list(AGG_FUNCS.keys())
        func.remove("sum")
        func.remove("std")
        return func

    return [func] if isinstance(func, str) else func


def get_length(sf: SheetFrame, by: list[str], funcs: list | dict) -> int:
    n = 1 if isinstance(funcs, dict) else len(funcs)

    if not by:
        return n

    return len(sf[by].drop_duplicates()) * n


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
