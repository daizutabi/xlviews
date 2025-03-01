from __future__ import annotations

from collections.abc import Callable, Hashable
from itertools import product
from typing import TYPE_CHECKING, TypeAlias

import pandas as pd
from pandas import DataFrame

from .palette import ColorPalette, MarkerPalette, PaletteStyle, get_palette

if TYPE_CHECKING:
    from collections.abc import Iterator
    from typing import Self

    from xlviews.chart.axes import Axes
    from xlviews.chart.series import Series

Label: TypeAlias = str | Callable[[dict[str, Hashable]], str]


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
        label: Label | None = None,
        marker: PaletteStyle | None = None,
        color: PaletteStyle | None = None,
        alpha: float | None = None,
        weight: float | None = None,
        size: int | None = None,
    ) -> Self:
        index = self.data.index.to_frame(index=False)
        marker_palette = get_palette(MarkerPalette, index, marker)
        color_palette = get_palette(ColorPalette, index, color)

        for key, s in zip(self.keys(), self.series_collection, strict=True):
            s.set(
                label=label and get_label(label, key),
                color=color_palette and color_palette[key],
                marker=marker_palette and marker_palette[key],
                alpha=alpha,
                weight=weight,
                size=size,
            )

        return self


def get_label(label: Label, key: dict[str, Hashable]) -> str:
    if isinstance(label, str):
        return label.format(**key)

    if callable(label):
        return label(key)

    msg = f"Invalid label: {label}"
    raise ValueError(msg)
