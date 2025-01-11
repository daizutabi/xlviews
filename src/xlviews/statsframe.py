from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
from pandas import DataFrame
from xlwings.constants import Direction

from xlviews.common import turn_off_screen_updating
from xlviews.config import rcParams
from xlviews.formula import AGG_FUNCS, aggregate
from xlviews.range import multirange
from xlviews.sheetframe import SheetFrame
from xlviews.style import set_font
from xlviews.utils import iter_columns

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator

    from numpy.typing import NDArray
    from xlwings import Range, Sheet


class GroupedRange:
    sf: SheetFrame
    by: list[str]
    grouped: dict[Hashable, list[list[int]]]

    def __init__(self, sf: SheetFrame, by: str | list[str] | None = None) -> None:
        self.sf = sf
        self.by = list(iter_columns(sf, by)) if by else []
        self.grouped = sf.groupby(self.by)

    def get_length(self, funcs: list[str] | dict[str, str]) -> int:
        n = 1 if isinstance(funcs, dict) else len(funcs)
        return len(self.grouped) * n

    def iter_row_ranges(
        self,
        column: str,
        offset: int = 0,
    ) -> Iterator[str | list[Range]]:
        column_index = self.sf.index(column)
        if not isinstance(column_index, int):
            raise NotImplementedError

        index_columns = self.sf.index_columns
        sheet = self.sf.sheet

        for row in self.grouped.values():
            if column in self.by:
                start = row[0][0]
                yield sheet.range(start, column_index).offset(offset).get_address()

            elif column in index_columns:
                yield ""

            else:
                yield get_column_ranges(sheet, row, column_index, offset)

    def iter_formulas(
        self,
        column: str,
        funcs: list[str] | dict[str, str],
        wrap: str | None = None,
        offset: int = 0,
        default: str = "median",
    ) -> Iterator[str]:
        for ranges in self.iter_row_ranges(column, offset):
            if isinstance(funcs, dict):
                funcs = [funcs.get(column, default)]

            for func in funcs:
                yield get_formula(func, ranges, wrap)

    def get_index(self, funcs: list[str]) -> list[str]:
        return funcs * len(self.grouped)

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
        offset: int = 0,
        default: str = "median",
    ) -> NDArray[np.str_]:
        values = [self.get_index(funcs)] if isinstance(funcs, list) else []

        for column in self.sf.columns:
            wrap_ = wrap.get(column) if isinstance(wrap, dict) else wrap
            it = self.iter_formulas(column, funcs, wrap_, offset, default)
            values.append(list(it))

        return np.array(values).T

    def get_frame(
        self,
        funcs: list[str] | dict[str, str],
        wrap: str | dict[str, str] | None = None,
        offset: int = 0,
        default: str = "median",
        func_column_name: str = "func",
    ) -> DataFrame:
        values = self.get_values(funcs, wrap, offset, default)
        columns = self.get_columns(funcs, func_column_name)
        df = DataFrame(values, columns=columns)
        return df.set_index(columns[: -len(self.sf.value_columns)])


def get_column_ranges(
    sheet: Sheet,
    row: list[list[int]],
    column: int,
    offset: int = 0,
) -> list[Range]:
    rngs = []

    for start, end in row:
        ref = sheet.range((start + offset, column), (end + offset, column))
        rngs.append(ref)

    return rngs


def get_formula(
    func: str | Range,
    ranges: str | list[Range],
    wrap: str | None = None,
) -> str:
    if not ranges:
        return ""

    if isinstance(ranges, str):
        return f"={ranges}"

    formula = aggregate(func, *ranges)

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
        """
        統計値シートフレームを作成する。

        Parameters
        ----------
        parent : SheetFrame
            集計対象シートフレーム
        funcs : str, list of str, dict, optional
            集計関数を指定する。指定できる関数は以下
                'count', 'sum', 'min', 'max', 'mean', 'median', 'std', 'soa%'
            Noneとするとデフォルト関数を使う。
        by : str, list of str, optional
            集計をグルーピングするカラム名
        autofilter : bool
            Trueのとき、表示される関数を所定のものに制限する。
        table : bool
            Trueのとき、エクセルの表形式にする。
        wrap : str or dict, optional
            統計関数をラップする文字列。format形式で{}が統計関数に
            置き換えられる。
        na : bool
            Trueのとき、self.wrap = 'IFERROR({},NA())'となる。
        null : bool
            Trueのとき、self.wrap = 'IFERROR({},"")'となる。
        succession : bool, optional
            連続インデックスを隠すか
        **kwargs
            SheetFrame.__init__関数に渡される。
        """
        wrap = get_wrap(wrap, na=na, null=null)
        funcs = get_func(funcs)
        gr = GroupedRange(parent, by)

        offset = gr.get_length(funcs) + 2

        row = parent.row
        column = parent.column
        if isinstance(funcs, list):
            column -= 1

        offset = move_down(parent, offset)
        df = gr.get_frame(funcs, wrap, offset, default, func_column_name)

        super().__init__(
            parent.sheet,
            row,
            column,
            data=df,
            index=parent.has_index,
            autofit=False,
            style=False,
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
        start = self.column + self.index_level
        end = self.column + len(self.columns)
        func_index = self.index(func_column_name)

        value_columns = ["median", "min", "mean", "max", "std", "sum"]

        grouped = self.groupby(func_column_name)
        columns = [func_index, *range(start, end)]
        get_fmt = self.parent.get_number_format
        formats = [get_fmt(column) for column in self.value_columns]
        formats = [None, *formats]

        for func, rows in grouped.items():
            for column, fmt in zip(columns, formats, strict=False):
                cell = multirange(self.sheet, rows, column)  # type: ignore
                if func in value_columns:
                    cell.number_format = fmt

                color = rcParams.get(f"stats.{func}.color")
                italic = rcParams.get(f"stats.{func}.italic")
                set_font(cell, color=color, italic=italic)

                if func == "soa" and column != func_index:
                    cell.number_format = "0.0%"

        set_font(self.range(func_column_name, -1), italic=True)


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
