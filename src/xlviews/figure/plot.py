from __future__ import annotations

from collections.abc import Callable, Hashable
from itertools import product
from typing import TYPE_CHECKING, TypeAlias

import pandas as pd
from pandas import DataFrame, Index

from .palette import PaletteStyle, get_color_palette, get_marker_palette

if TYPE_CHECKING:
    from collections.abc import Iterator, Sequence
    from typing import Any, Self

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
        marker_palette = get_marker_palette(self.data, marker)
        color_palette = get_color_palette(self.data, color)

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

    @classmethod
    def facet(
        cls,
        axes: Axes,
        data: DataFrame,
        index: str | list[str] | None = None,
        columns: str | list[str] | None = None,
    ) -> Iterator[tuple[dict[str, Any], Self]]:
        left = axes.chart.left
        top = axes.chart.top
        width = axes.chart.width
        height = axes.chart.height

        for r, rkey in enumerate(iterrows(data.index, index)):
            for c, ckey in enumerate(iterrows(data.index, columns)):
                key = rkey | ckey
                sub = xs(data, key)

                if len(sub) == 0:
                    continue

                if r == 0 and c == 0:
                    axes_ = axes
                else:
                    axes_ = axes.copy(left=left + c * width, top=top + r * height)

                yield key, cls(axes_, sub)


def get_label(label: Label, key: dict[str, Hashable]) -> str:
    if isinstance(label, str):
        return label.format(**key)

    if callable(label):
        return label(key)

    msg = f"Invalid label: {label}"
    raise ValueError(msg)


def iterrows(
    index: Index,
    levels: int | str | Sequence[int | str] | None,
) -> Iterator[dict[str, Any]]:
    if levels is None:
        yield {}
        return

    if isinstance(levels, int | str):
        levels = [levels]

    if levels:
        values = {level: index.get_level_values(level) for level in levels}
        it = DataFrame(values).drop_duplicates().iterrows()

        for _, s in it:
            yield s.to_dict()


def xs(df: DataFrame, index: dict[str, Any] | None) -> DataFrame:
    if index:
        df = df.xs(tuple(index.values()), 0, tuple(index.keys()), drop_level=False)  # type: ignore

    return df
