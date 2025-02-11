from __future__ import annotations

from itertools import product

import numpy as np
import pandas as pd
from pandas import DataFrame

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.range import Range
from xlviews.range.style import set_alignment, set_font
from xlviews.utils import iter_columns

from .sheet_frame import SheetFrame


class DistFrame(SheetFrame):
    dist_func: dict[str, str]

    @turn_off_screen_updating
    def __init__(
        self,
        parent: SheetFrame,
        columns: str | list[str] | None = None,
        by: str | list[str] | None = None,
        dist: str | dict[str, str] = "norm",
    ) -> None:
        if columns is None:
            columns = parent.value_columns
        elif isinstance(columns, str):
            columns = [columns]

        self.dist_func = get_dist_func(dist, columns)

        if not parent.table:
            parent.as_table()

        by = list(iter_columns(parent, by)) if by else None
        data = get_init_data(parent, columns, by)

        row = parent.row
        column = parent.column + len(parent.columns) + 1

        super().__init__(row, column, data, index=by is not None, sheet=parent.sheet)

        if by:
            self.link_to_index(parent, by)

        self.set_values(parent, columns, by)
        self.style(gray=True)
        self.autofit()
        self.const_values(parent)

    def set_values(
        self,
        parent: SheetFrame,
        columns: list[str],
        by: list[str] | None,
    ) -> None:
        group = self.groupby(by)
        parent_columns = parent.column_index(columns)
        self_columns = self.column_index([c + "_n" for c in columns])

        it = zip(parent_columns, self_columns, columns, strict=True)
        for parent_column, self_column, column in it:
            dist = self.dist_func[column]

            for row in group.values():
                if len(row) != 1:
                    raise ValueError("group must be continuous")

                start = row[0][0]
                length = row[0][1] - start + 1

                parent_cell = Range((start, parent_column), sheet=parent.sheet)
                self_cell = Range((start, self_column), sheet=self.sheet)

                formula = counter(parent_cell)
                set_formula(self_cell, length, formula)

                formula = sorted_value(parent_cell, self_cell, length)
                set_formula(self_cell.offset(0, 1), length, formula)

                formula = sigma_value(self_cell, length, dist)
                set_formula(self_cell.offset(0, 2), length, formula)

        for column in columns:
            fmt = parent.range(column).impl.number_format
            self.number_format({f"{column}_v": fmt, f"{column}_s": "0.00"})

    def link_to_index(self, parent: SheetFrame, by: list[str]) -> None:
        start = self.row + 1
        end = start + len(self) - 1

        parent_columns = parent.column_index(by)
        self_columns = self.column_index(by)

        it = zip(parent_columns, self_columns, strict=True)
        for parent_column, self_column in it:
            ref = Range((start, parent_column), sheet=parent.sheet)
            to = Range((start, self_column), (end, self_column), sheet=self.sheet)
            to.value = ref.get_address(row_absolute=False, formula=True)

    def const_values(self, parent: SheetFrame) -> None:
        index = parent.index_columns
        array = np.zeros((len(index), 1))
        data = DataFrame(array, columns=["value"], index=index)

        row = self.row + len(self) + self.columns_level + 1
        column = self.column

        sf = SheetFrame(row, column, data, index=bool(index), sheet=self.sheet)
        sf.style(gray=True)

        head = Range((parent.row - 1, parent.column), sheet=parent.sheet)
        tail = Range((sf.row + 1, sf.column + 1), sheet=sf.sheet)

        for k in range(len(index)):
            tail.offset(k, 0).value = head.offset(0, k).get_address(formula=True)

    def plot(
        self,
        x,
        label="auto",
        color=None,
        marker=None,
        axes=None,
        xlabel="auto",
        ylabel="auto",
        **kwargs,
    ):
        if ylabel == "auto":
            dist = self.dist_func[x] if isinstance(x, str) else self.dist_func[x[0]]
            ylabel = "σ" if dist == "norm" else "ln(-ln(1-F))"
        plot = None
        if isinstance(x, str) and xlabel == "auto":
            x_ = x.split("_")[0]
            xlabel = rcParams.get(f"axis.label.{x_}", x)
            if "_" in x and "[" in xlabel:
                xlabel = x + " " + xlabel[xlabel.index("[") :]
        xs = [x] if isinstance(x, str) else x
        colors = color if isinstance(color, list) else [color] * len(xs)
        markers = marker if isinstance(marker, list) else [marker] * len(xs)
        for x_, color, marker in zip(xs, colors, markers, strict=False):
            label_ = x_ if label == "auto" and isinstance(x, list) else label
            plot = self._plot(
                x_,
                label=label_,
                axes=axes,
                color=color,
                marker=marker,
                xlabel=xlabel,
                ylabel=ylabel,
                **kwargs,
            )
            axes = plot.axes
        return plot

    def _plot(self, x, **kwargs):
        plot = super().plot(f"{x}_v", f"{x}_s", yformat="0_ ", **kwargs)
        # if fit:
        #     sigma = 2 if fit is True else fit
        #     column = self.add_column_for_fit(x, sigma)
        #     kwargs['marker'] = None
        #     kwargs['line'] = None
        #     kwargs.pop('axes')
        #     kwargs.pop('label')
        #     plot_ = super().plot(f'{x}_v', column, axes=plot.axes,
        #                          label=None, **kwargs)
        #     for series in plot_.series_collection:
        #         trendline = series.Trendlines().Add()
        #         plot_.axes.labels.append('__trendline__')
        #         trendline.DisplayEquation = True
        #         # trendline.Forward = 10
        #         # trendline.Backward = 10
        #     # print(plot_.axes.labels)
        #     # print(plot_.axes.legend.LegendEntries())
        #     # print(plot_.legend)
        #     plot_.axes.set_legend(**plot_.legend)
        return plot

    def fit(self, x):
        pass

    def add_column_for_fit(self, x, sigma):
        """

        Parameters
        ----------
        x : str
            変数名
        sigma: int or float
            フィッティングに用いるσ値の範囲
        """
        column_ = f"{x}_sf"
        if column_ in self.columns:
            return column_
        self[column_] = 1
        column = self.index(column_)
        sigma_cell = self.sheet.range(self.row - 1, column)
        sigma_cell.value = sigma
        set_font(sigma_cell, size=8, bold=True, italic=True, color="green")
        set_alignment(sigma_cell, "center")
        sigma = sigma_cell.get_address()
        row = self.cell.offset(self.columns_level).row
        column_ref = self.index(f"{x}_s")
        cell_ref = self.sheet.range(row, column_ref)
        cell_ref = cell_ref.get_address(row_absolute=False)
        cell = self.sheet.range(row, column)
        range_ = self.sheet.range(cell, cell.offset(len(self) - 1))
        range_.api.NumberFormatLocal = "0.00_ "
        formula = f"=IF(AND({cell_ref}>=-{sigma},{cell_ref}<={sigma}),{cell_ref},NA())"
        range_.value = formula
        return column_


def get_init_data(
    sf: SheetFrame,
    columns: list[str],
    by: list[str] | None,
) -> DataFrame:
    columns = [f"{c}_{suffix}" for c, suffix in product(columns, ["n", "v", "s"])]

    array = np.zeros((len(sf), len(columns)))
    df = DataFrame(array, columns=columns)

    if by:
        index = np.zeros((len(sf), len(by)))
        index = DataFrame(index, columns=by)
        df = pd.concat([index, df], axis=1)
        df = df.set_index(by)

    return df


def get_dist_func(dist: str | dict[str, str], columns: list[str]) -> dict[str, str]:
    if isinstance(dist, str):
        return {column: dist for column in columns}

    dist = dist.copy()
    for column in columns:
        dist.setdefault(column, "norm")

    return dist


def counter(cell: Range) -> str:
    start = cell.get_address()
    end = cell.get_address(row_absolute=False)
    return f"=AGGREGATE(3,1,{start}:{end})"


def sorted_value(parent_cell: Range, cell: Range, length: int) -> str:
    start = parent_cell.get_address()
    end = parent_cell.offset(length - 1).get_address()
    small = cell.get_address(row_absolute=False)
    return f"=IF({small}>0,AGGREGATE(15,1,{start}:{end},{small}),NA())"


def sigma_value(cell: Range, length: int, dist: str) -> str:
    small = cell.get_address(row_absolute=False)
    end = cell.offset(length - 1).get_address()

    if dist == "norm":
        return f"=IF({small}>0,NORM.S.INV({small}/({end}+1)),NA())"

    if dist == "weibull":
        return f"=IF({small}>0,LN(-LN(1-{small}/({end}+1))),NA())"

    msg = f"unknown distribution: {dist}"
    raise ValueError(msg)


def set_formula(cell: Range, length: int, formula: str) -> None:
    rng = Range((cell.row, cell.column), (cell.row + length - 1, cell.column))
    rng.value = formula
