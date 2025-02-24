from __future__ import annotations

from itertools import product
from typing import TYPE_CHECKING, Any, TypeAlias, overload

import pandas as pd
from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.dataframes.sheet_frame import SheetFrame

from .style import Label, Style, format_label, format_style

if TYPE_CHECKING:
    from collections.abc import Callable, Hashable, Iterable, Iterator
    from typing import Self

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy


def plot_series(
    ax: Axes,
    series: pd.Series,
    x: str,
    y: str,
    label: Label = None,
    marker: Style = None,
    color: Style = None,
    alpha: float | None = None,
    weight: float | None = None,
    size: int | None = None,
) -> Series:
    label = format_label(label, series.name)
    marker = format_style(marker, series.name)
    color = format_style(color, series.name)
    s = ax.add_series(series[x], series[y], label=label)
    return s.set(marker=marker, color=color, alpha=alpha, weight=weight, size=size)


def plot(
    ax: Axes,
    data: DataFrame | pd.Series,
    x: str,
    y: str,
    label: Label = None,
    marker: Style = None,
    color: Style = None,
    alpha: float | None = None,
    weight: float | None = None,
    size: int | None = None,
) -> Series:
    if isinstance(data, pd.Series):
        return plot_series(ax, data, x, y, label, marker, color, alpha, weight, size)

    names = data.index.names

    for key, s in data.iterrows():
        print(key, s)


class Plot:
    axes: Axes
    data: DataFrame
    keys: list[Hashable]
    series_collection: list[Series]

    def __init__(self, axes: Axes, data: DataFrame | pd.Series) -> None:
        self.axes = axes

        if isinstance(data, pd.Series):
            data = data.to_frame().T

        self.data = data

    def add(
        self,
        x: str | list[str],
        y: str | list[str],
        chart_type: int | None = None,
    ) -> Self:
        self.keys = []
        self.series_collection = []

        xs = x if isinstance(x, list) else [x]
        ys = y if isinstance(y, list) else [y]

        for x, y in product(xs, ys):
            for key, series in self.data.iterrows():
                s = self.axes.add_series(series[x], series[y], chart_type=chart_type)
                self.keys.append(key)
                self.series_collection.append(s)

        return self
