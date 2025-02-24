from __future__ import annotations

from itertools import product
from typing import TYPE_CHECKING, Any, TypeAlias, overload

import pandas as pd
from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.dataframes.sheet_frame import SheetFrame

from .style import Label, format_label

if TYPE_CHECKING:
    from collections.abc import Callable, Hashable, Iterable, Iterator
    from typing import Self

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy


class Plot:
    axes: Axes
    data: DataFrame
    index: list[tuple[Hashable, ...]]
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
        self.index = []
        self.series_collection = []

        xs = x if isinstance(x, list) else [x]
        ys = y if isinstance(y, list) else [y]

        for x, y in product(xs, ys):
            for idx, s in self.data.iterrows():
                series = self.axes.add_series(s[x], s[y], chart_type=chart_type)
                index = idx if isinstance(idx, tuple) else (idx,)
                self.index.append(index)
                self.series_collection.append(series)

        return self

    def keys(self) -> Iterator[dict[str, Hashable]]:
        names = self.data.index.names

        for index in self.index:
            yield dict(zip(names, index, strict=True))

    def set(
        self,
        label: Label = None,
        # marker: Style = None,
        # color: Style = None,
        # alpha: float | None = None,
        # weight: float | None = None,
        # size: int | None = None,
    ) -> Self:
        for key, s in zip(self.keys(), self.series_collection, strict=True):
            print(key)
            # label_ = format_label(label, key)
            # marker_ = format_style(marker, key)
            # color_ = format_style(color, key)

            s.set(
                label=format_label(label, key),
                # marker=marker_,
                # color=color_,
                # alpha=alpha,
                # weight=weight,
                # size=size,
            )

        return self
