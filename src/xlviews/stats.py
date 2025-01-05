from collections import OrderedDict

import numpy as np
import pandas as pd
import xlwings as xw

from xlviews.decorators import wait_updating
from xlviews.formula import AGG_FUNCS, aggregate
from xlviews.frame import SheetFrame
from xlviews.range import multirange
from xlviews.style import set_alignment, set_font
from xlviews.utils import constant, iter_columns, outline_group, outline_levels


class StatsFrame(SheetFrame):
    @wait_updating
    def __init__(
        self,
        parent,
        by=None,
        stats=None,
        group=False,
        link=False,
        autofilter=True,
        table=True,
        wrap=None,
        na=False,
        null=False,
        default="median",
        succession=False,
        **kwargs,
    ):
        """
        統計値シートフレームを作成する。

        Parameters
        ----------
        parent : SheetFrame
            集計対象シートフレーム
        by : str, list of str, optional
            集計をグルーピングするカラム名
        stats : str, list of str, dict, optional
            集計関数を指定する。指定できる関数は以下：
                'count', 'sum', 'min', 'max', 'mean', 'median', 'std', 'soa%'
            Noneとするとデフォルト関数を使う。
        group : bool
            エクセルのグループ機能を使うかどうか
        link : bool
            Trueのとき、統計関数をセル参照で表現する。
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
        default : str or None
            statsが辞書のときに指定されないカラムのデフォルト関数
        succession : bool, optional
            連続インデックスを隠すか
        **kwargs
            SheetFrame.__init__関数に渡される。
        """
        self.by = list(iter_columns(parent, by)) if by else None
        self.func_column = "func"
        self.func = AGG_FUNCS.copy()
        self.link = link
        self.grouped = None
        self.length = 1
        self.wrap = None

        if na is True:
            self.wrap = "IFERROR({},NA())"
        elif null is True:
            self.wrap = 'IFERROR({},"")'
        elif wrap:
            self.wrap = wrap
        else:
            self.wrap = {}
            if na:
                nas = [na] if isinstance(na, str) else na  # type: list
                for na in nas:
                    self.wrap[na] = "IFERROR({},NA())"
            if null:
                nulls = [null] if isinstance(null, str) else null  # type: list
                for null in nulls:
                    self.wrap[null] = 'IFERROR({},"")'

        self.default = default

        if stats:
            if isinstance(stats, str):
                stats = [stats]
            if isinstance(stats, dict):
                self.func = {"__dummy__": -1}
                stats_ = {}
                for key, value in stats.items():
                    if key not in parent.value_columns:
                        for column in parent.range(key, 0):
                            stats_[column.value] = value
                    else:
                        stats_[key] = value
                stats = stats_
            else:
                for key in list(self.func.keys()):
                    if key not in stats:
                        self.func.pop(key)
            self.column_func = stats
        else:
            self.func.pop("sum")
            self.func.pop("std")
            self.column_func = list(self.func.keys())

        df = self.dummy_dataframe(parent)

        start_ = parent.row
        parent_header = parent.sheet.range(
            parent.cell.offset(-1),
            parent.cell.offset(-1, len(parent.columns)),
        )
        if any(parent_header.value):
            start = start_ - 1
            end = start + len(df) + 2
            start_ = start + 1
        else:
            start = start_
            end = start + len(df) + 1

        sheet = parent.sheet
        rows = parent.sheet.api.Rows(f"{start}:{end}")
        rows.Insert(Shift=constant("Direction.xlDown"))

        column = parent.column
        if not isinstance(self.column_func, dict):
            column -= 1

        super().__init__(
            sheet,
            start_,
            column,
            data=df,
            index=parent.index_level > 0,
            autofit=False,
            style=False,
            **kwargs,
        )
        self.parent = parent

        # オフセット
        if self.grouped:
            row_offset = len(list(rows))
            for key, value in self.grouped.items():
                value = [[v + row_offset for v in vv] for vv in value]
                self.grouped[key] = value

        self._aggregate()

        if table:
            self.as_table(autofit=False, const_header=True)

        # 罫線を正しく表示させるために、テーブル化の後にスタイル設定をする。
        self.set_style(autofit=False, succession=succession)

        if not isinstance(self.column_func, dict):
            self.set_value_style()

        if table:
            self.set_alignment("left")

        if table and autofilter and not isinstance(self.column_func, dict):
            if stats is None:
                func = "median"
            else:
                func = stats[0]
            self.autofilter(func=func)

        if group and len(self.func) > 1:
            group = group
            if group is True and not stats:
                group = [1, 2]
            self.group(group, level=1)

        # wide_columns = []
        # for k in range(len(parent.value_columns)):
        #     ref = parent.cell.offset(-1, parent.index_level + k)
        #     if ref.value:
        #         wide_columns.append(ref.value)
        #         header = self.cell.offset(-1, self.index_level + k)
        #         header.value = "=" + ref.get_address()
        #         set_font(header, bold=True, color="#002255")
        #         set_alignment(header, horizontal_alignment="center")
        # for column in wide_columns:
        #     ref = parent.range(column, 0)[0]
        #     column = self.range(column, 0)
        #     column.number_format = ref.number_format

        # print(wide_columns)
        # if wide_columns:
        #     self.set_style()

    def dummy_dataframe(self, parent):
        if self.by:
            self.grouped = OrderedDict()
            self.grouped.update(parent.groupby(self.by))
            self.length = len(self.grouped)

        if not isinstance(self.column_func, dict):
            columns = [self.func_column] + parent.columns
            array = np.zeros((self.length * len(self.func), len(columns)))
            df = pd.DataFrame(array, columns=columns)
            df[self.func_column] = [
                agg for _ in range(self.length) for agg in self.func
            ]
            if parent.index_level:
                df.set_index([self.func_column] + parent.index_columns, inplace=True)
            return df
        columns = parent.columns
        array = np.zeros((self.length, len(columns)))
        df = pd.DataFrame(array, columns=columns)
        if parent.index_level:
            df.set_index(parent.index_columns, inplace=True)
        return df

    def get_group_start_and_end(self):
        if self.by:
            for _, row in self.grouped.items():
                yield row
        else:
            start = self.parent.row + 1
            end = start + len(self.parent) - 1
            yield [[start, end]]

    def _aggregate(self):
        by_ = self.by if self.by else []
        parent_columns = self.parent.columns
        parent_index_columns = self.parent.index_columns
        values = np.empty(
            (self.length * len(self.func), len(parent_columns)),
            dtype=str,
        )
        values = pd.DataFrame(values)
        start_end = list(self.get_group_start_and_end())
        if not isinstance(self.column_func, dict):
            func_index = self.index(self.func_column)
        else:
            func_index = None

        for j, column in enumerate(parent_columns):
            column_index = self.parent.index(column)
            for i, start_end_ in enumerate(start_end):
                start = start_end_[0][0]
                if column in by_:
                    ref = self.sheet.range(start, column_index)
                    ref = ref.get_address()
                    formula = f"{ref}"
                else:
                    if column in parent_index_columns:
                        continue
                        # formula = const(ref)
                    refs = []
                    for start, end in start_end_:
                        ref = self.sheet.range(
                            (start, column_index),
                            (end, column_index),
                        )
                        refs.append(ref)
                    formula = refs
                for r, func in enumerate(self.func):
                    index = i * len(self.func) + r
                    row = index + self.row + 1
                    if isinstance(formula, str):
                        formula_ = formula
                    else:
                        if isinstance(self.column_func, dict):
                            func = self.column_func.get(column, self.default)
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
        column_offset = 0 if isinstance(self.column_func, dict) else 1
        self.cell.offset(1, column_offset).value = values.values

    def group(self, group, level=1):
        for i in range(self.length):
            end = self.row + (i + 1) * len(self.func)
            for g in group:
                start = self.row + i * len(self.func) + g + 1
                outline_group(self.sheet, start, end)
        if level:
            outline_levels(self.sheet, level)

    def set_value_style(self):
        start = self.column + self.index_level
        end = self.column + len(self.columns)
        func_index = self.index(self.func_column)
        value_columns = ["median", "min", "mean", "max", "std", "sum"]
        grouped = self.groupby("func")
        columns = [func_index] + list(range(start, end))
        formats = [
            self.parent.get_number_format(column) for column in self.value_columns
        ]
        formats = [None] + formats
        for func, rows in grouped.items():
            for column, format_ in zip(columns, formats, strict=False):
                cell = multirange(self.sheet, rows, column)
                if func in value_columns:
                    set_number_format(cell, format_)
                if func == "soa":
                    if column != func_index:
                        set_number_format(cell, "0.0%")
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
        set_font(self.range("func", -1), italic=True)


def main() -> None:
    from pandas import DataFrame

    from xlviews.common import quit_apps
    from xlviews.style import hide_gridlines

    quit_apps()

    book = xw.Book()
    sheet = book.sheets[0]
    hide_gridlines(sheet)

    df = DataFrame(
        {
            "x": ["a"] * 10 + ["b"] * 10,
            "y": (["c"] * 6 + ["d"] * 4) * 2,
            "z": range(1, 21),
            "a": range(20),
            "b": list(range(10)) + list(range(0, 30, 3)),
            "c": list(range(20, 40, 2)) + list(range(0, 20, 2)),
        },
    )
    df = df.set_index(["x", "y", "z"])
    df.iloc[[4, -1], 0] = np.nan
    df.iloc[[3, 6, 9], -1] = np.nan

    sf = SheetFrame(sheet, 2, 3, data=df, table=True)
    # sf.as_table()
    StatsFrame(sf, by=":y", stats={"a": "max", "b": "std", "c": "mean"}, table=True)
    # StatsFrame(sf, stats={"a": "max", "b": "min", "c": "mean"})
    from scipy.stats import norm

    print(norm.std([1, 2, 3]))

    # directory = mtj.get_directory("local", "Data")
    # run = mtj.get_paths_dataframe(directory, "S6544", "IB01-06")
    # series = run.iloc[0]
    # path = mtj.get_path(directory, series)
    # with mtj.data(path) as data:
    #     data.merge_device()
    #     df = data.get(
    #         ["wafer", "cad", "sx", "sy", "dx", "dy", "id", "Rmin", "Rmax", "TMR"],
    #         sx=(4, 6),
    #         sy=(3, 4),
    #     )
    #     df.set_index(["wafer", "cad", "sx", "sy", "dx", "dy", "id"], inplace=True)
    #     sf = xv.SheetFrame(sheet, 2, 3, data=df, sort_index=True)
    #     sf.set_number_format(Rmin="0.00", TMR="0.0")
    # sf.astable()

    # start_time = time.time()
    # StatsFrame(sf, by=":sy", stats={"Rmin": "max", "Rmax": "count"}, default=None)
    # # , autofilter=True)
    # # sf.autofilter('func')
    # # sf.autofilter(func='median')
    # elapsed_time = time.time() - start_time
    # print(f"elapsed_time:{elapsed_time}" + "[sec]")


if __name__ == "__main__":
    main()
