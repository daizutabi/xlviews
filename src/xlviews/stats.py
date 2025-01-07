from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
import xlwings as xw
from pandas import DataFrame
from xlwings.constants import Direction

from xlviews.decorators import wait_updating
from xlviews.formula import AGG_FUNCS, aggregate
from xlviews.frame import SheetFrame
from xlviews.range import multirange
from xlviews.style import set_alignment, set_font
from xlviews.table import Table
from xlviews.utils import constant, iter_columns, outline_group, outline_levels

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator

    from xlwings import Range, Sheet


class GroupedRange:
    sf: SheetFrame
    by: list[str]
    grouped: dict[Hashable, list[list[int]]]

    def __init__(self, sf: SheetFrame, by: str | list[str] | None = None) -> None:
        self.sf = sf
        self.by = list(iter_columns(sf, by)) if by else []
        self.grouped = sf.groupby(self.by)

    def iter_row_ranges(self, column: str) -> Iterator[str | list[Range]]:
        column_index = self.sf.index(column)
        if not isinstance(column_index, int):
            raise NotImplementedError

        index_columns = self.sf.index_columns
        sheet = self.sf.sheet

        for row in self.grouped.values():
            if column in self.by:
                start = row[0][0]
                yield sheet.range(start, column_index).get_address()

            elif column in index_columns:
                yield ""

            else:
                yield get_column_ranges(sheet, row, column_index)

    def iter_formulas(
        self,
        column: str,
        funcs: list[str] | dict[str, str],
        wrap: str | None = None,
        default: str = "median",
    ) -> Iterator[str]:
        for ranges in self.iter_row_ranges(column):
            if isinstance(funcs, dict):
                funcs = [funcs.get(column, default)]

            for func in funcs:
                yield get_formula(func, ranges, wrap)

    def get_index(self, funcs: list[str]) -> list[str]:
        return funcs * len(self.grouped)


def get_column_ranges(sheet: Sheet, row: list[list[int]], column: int) -> list[Range]:
    rngs = []

    for start, end in row:
        ref = sheet.range((start, column), (end, column))
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

    if not (formula := aggregate(func, *ranges)):
        return ""

    if wrap:
        formula = wrap.format(formula)

    if formula:
        return "=" + formula

    return ""


class StatsFrame(SheetFrame):
    parent: SheetFrame
    funcs: list[str] | dict[str, str]
    by: list[str]
    wrap: str | dict[str, str]
    grouped: dict[Hashable, list[list[int]]]
    func_index_name: str = "func"
    default: str = "median"

    @wait_updating
    def __init__(
        self,
        parent: SheetFrame,
        funcs: str | list[str] | dict[str, str] | None = None,
        *,
        by: str | list[str] | None = None,
        autofilter: bool = True,
        table: bool = True,
        wrap: str | dict[str, str] | None = None,
        na: str | list[str] | bool = False,
        null: str | list[str] | bool = False,
        succession: bool = False,
        # group=False,
        **kwargs,
    ) -> None:
        """
        統計値シートフレームを作成する。

        Parameters
        ----------
        parent : SheetFrame
            集計対象シートフレーム
        funcs : str, list of str, dict, optional
            集計関数を指定する。指定できる関数は以下：
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
        self.wrap = get_wrap(wrap, na=na, null=null)
        self.funcs = get_func(funcs)

        self.by = list(iter_columns(parent, by)) if by else []
        self.grouped = parent.groupby(self.by)

        groups = len(self.grouped)
        df = get_init_data(parent, self.funcs, groups, self.func_index_name)

        row = parent.row
        column = parent.column
        if isinstance(self.funcs, list):
            column -= 1

        row_offset = move_down(parent, len(df) + 2)

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

        if self.grouped:
            for key, value in self.grouped.items():
                self.grouped[key] = [[x + row_offset for x in v] for v in value]

        self.aggregate()

        return

        if table:
            self.as_table(autofit=False, const_header=True)

        # 罫線を正しく表示させるために、テーブル化の後にスタイル設定をする。
        self.set_style(autofit=False, succession=succession)

        if not isinstance(self.funcs, dict):
            self.set_value_style()

        if table:
            self.set_alignment("left")

        if self.table and autofilter and not isinstance(self.funcs, dict):
            if stats is None:
                funcs = "median"
            else:
                funcs = stats[0]
            self.table.auto_filter(func=funcs)

    def iter_rows(self, column: str) -> Iterator[str]:
        if self.link:  # noqa: SIM108
            func_index = self.index(self.func_index_name)
        else:
            func_index = None

        if isinstance(self.wrap, dict):  # noqa: SIM108
            wrap = self.wrap.get(column, "{}")
        else:
            wrap = self.wrap

        for k, ranges in enumerate(self.iter_row_ranges(column)):
            if isinstance(self.funcs, dict):
                funcs = [self.funcs.get(column, self.default)]
            elif self.link:
                n = len(self.funcs)
                row = self.row + k * n
                funcs = [self.sheet.range(row + i, func_index) for i in range(n)]
            else:
                funcs = self.funcs

            for func in funcs:
                yield get_formula(func, ranges, self.wrap, column)

            # else:
            #     funcs = self.func

            # formula_ = aggregate(func, *formula)
            # if self.wrap and formula_:
            #     if isinstance(self.wrap, dict):
            #         wrap = self.wrap.get(column, "{}")

            # if isinstance(self.func, dict):
            #     if isinstance(ranges, str):
            #         yield ranges
            #     else:
            #         for range in ranges:
            #             yield range.get_address()
            # else:
            #     yield ranges

    def aggregate(self):
        parent_columns = self.parent.columns
        parent_index_columns = self.parent.index_columns

        groups = len(self.grouped) or 1
        values = get_empty_data(self.parent, self.funcs, groups)

        if not isinstance(self.funcs, dict):
            func_index = self.index(self.func_index_name)
        else:
            func_index = None

        for j, column in enumerate(parent_columns):
            column_index = self.parent.index(column)

            for i, start_end in enumerate(self.grouped.values()):
                if column in self.by:
                    start = start_end[0][0]
                    formula = self.sheet.range(start, column_index).get_address()
                elif column in parent_index_columns:
                    continue
                else:
                    formula = get_column_ranges(self.sheet, start_end, column_index)

                if isinstance(self.funcs, list):
                    for r, func in enumerate(self.funcs):
                        index = i * len(self.funcs) + r
                        row = index + self.row + 1

                        if isinstance(formula, str):
                            formula_ = formula
                        else:
                            if isinstance(self.funcs, dict):
                                func = self.funcs.get(column, self.default)
                            elif self.link:
                                func = self.sheet.range(row, func_index)
                            formula_ = aggregate(func, *formula)
                            if self.wrap and formula_:
                                if isinstance(self.wrap, dict):
                                    wrap = self.wrap.get(column, "{}")
                                else:
                                    wrap = self.wrap
                                formula_ = wrap.format(formula_)
                        if formula_:
                            values.loc[index, j] = "=" + formula_
                else:
                    index = i * len(self.funcs) + r
                    row = index + self.row + 1

                    if isinstance(formula, str):
                        formula_ = formula
                    else:
                        if isinstance(self.funcs, dict):
                            func = self.funcs.get(column, self.default)
                        elif self.link:
                            func = self.sheet.range(row, func_index)
                        formula_ = aggregate(func, *formula)
                        if self.wrap and formula_:
                            if isinstance(self.wrap, dict):
                                wrap = self.wrap.get(column, "{}")
                            else:
                                wrap = self.wrap
                            formula_ = wrap.format(formula_)
                    if formula_:
                        values.loc[index, j] = "=" + formula_

            column_offset = 0 if isinstance(self.funcs, dict) else 1
        self.cell.offset(1, column_offset).value = values.values

    def set_value_style(self):
        start = self.column + self.index_level
        end = self.column + len(self.columns)
        func_index = self.index(self.func_index_name)
        value_columns = ["median", "min", "mean", "max", "std", "sum"]
        grouped = self.groupby(self.func_index_name)
        columns = [func_index] + list(range(start, end))
        formats = [
            self.parent.get_number_format(column) for column in self.value_columns
        ]
        formats = [None] + formats
        for func, rows in grouped.items():
            for column, format_ in zip(columns, formats, strict=False):
                cell = multirange(self.sheet, rows, column)
                if func in value_columns:
                    cell.number_format = format_
                if func == "soa":
                    if column != func_index:
                        cell.number_format = "0.0%"
                    set_font(cell, color="#5555FF", italic=True)
                elif func == "min":
                    set_font(cell, color="#7777FF")
                elif func == "mean":
                    set_font(cell, color="#33aa33")
                elif func == "max":
                    set_font(cell, color="#FF7777")
                elif func == "sum":
                    set_font(cell, color="purple", italic=True)
                elif func == "std":
                    set_font(cell, color="#aaaaaa")
                elif func == "count":
                    set_font(cell, color="gray")
        set_font(self.range(self.func_index_name, -1), italic=True)

    def auto_filter(self, func: str | list[str]) -> None:
        if self.table:
            self.table.auto_filter(self.func_index_name, func)

    # def group(self, group, level=1):
    #     for i in range(self.length):
    #         end = self.row + (i + 1) * len(self.func)
    #         for g in group:
    #             start = self.row + i * len(self.func) + g + 1
    #             outline_group(self.sheet, start, end)
    #     if level:
    #         outline_levels(self.sheet, level)


def get_wrap(
    wrap: str | dict[str, str] | None = None,
    *,
    na: str | list[str] | bool = False,
    null: str | list[str] | bool = False,
) -> str | dict[str, str]:
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

    return wrap


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
    start = sf.row
    end = sf.row + length - 1

    if has_header(sf):
        start -= 1

    rows = sf.sheet.api.Rows(f"{start}:{end}")
    rows.Insert(Shift=Direction.xlDown)
    return len(list(rows))


def get_init_data(
    sf: SheetFrame,
    func: list | dict,
    groups: int,
    func_index_name: str,
) -> DataFrame:
    if isinstance(func, dict):
        columns = sf.columns
        array = np.zeros((groups, len(columns)))

        df = DataFrame(array, columns=columns)

        if sf.index_level:
            df = df.set_index(sf.index_columns)

        return df

    columns = [func_index_name, *sf.columns]
    array = np.zeros((groups * len(func), len(columns)))

    df = DataFrame(array, columns=columns)
    df[func_index_name] = [f for _ in range(groups) for f in func]

    if sf.index_level:
        df = df.set_index([func_index_name, *sf.index_columns])

    return df


if __name__ == "__main__":
    import xlwings as xw

    from xlviews.common import quit_apps
    from xlviews.frame import SheetFrame

    df = DataFrame(
        {
            "x": ["a"] * 8 + ["b"] * 8 + ["a"] * 4,
            "y": (["c"] * 4 + ["d"] * 4) * 2 + ["c"] * 4,
            "z": range(1, 21),
            "a": range(20),
            "b": list(range(10)) + list(range(0, 30, 3)),
            "c": list(range(20, 40, 2)) + list(range(0, 20, 2)),
        },
    )
    df = df.set_index(["x", "y", "z"])
    df.iloc[[4, -1], 0] = np.nan
    df.iloc[[3, 6, 9], -1] = np.nan

    quit_apps()
    book = xw.Book()
    sheet = book.sheets.add()
    sf = SheetFrame(sheet, 2, 3, data=df, table=True)
    # sf = StatsFrame(
    #     sf,
    #     by="x",
    #     stats={"a": "count", "b": "median", "c": "std"},
    #     table=True,
    # )
