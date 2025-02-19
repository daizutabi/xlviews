from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
from pandas import DataFrame, Index, MultiIndex

from xlviews.core.formula import aggregate
from xlviews.core.range import Range
from xlviews.dataframes.colorbar import Colorbar
from xlviews.style import set_color_scale
from xlviews.utils import suspend_screen_updates

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator
    from typing import Any, Self

    from numpy.typing import NDArray
    from pandas import Index
    from xlwings import Sheet


class HeatFrame(SheetFrame):
    index: Index
    columns: Index
    range: Range

    @suspend_screen_updates
    def __init__(
        self,
        row: int,
        column: int,
        data: DataFrame,
        sheet: Sheet | None = None,
        vmin: float | str | Range | None = None,
        vmax: float | str | Range | None = None,
    ) -> None:
        data = clean_data(data)

        super().__init__(row, column, data, sheet)

        self.columns = data.columns

        start = self.row + 1, self.column + 1
        end = start[0] + self.shape[0] - 1, start[1] + self.shape[1] - 1
        self.range = Range(start, end, self.sheet)

        set_heat_frame_style(self)
        self.set(vmin, vmax)

    def set(
        self,
        vmin: float | str | Range | None = None,
        vmax: float | str | Range | None = None,
    ) -> Self:
        rng = self.range

        if vmin is None:
            vmin = aggregate("min", rng)
        if vmax is None:
            vmax = aggregate("max", rng)

        set_color_scale(rng, vmin, vmax)
        return self

    def colorbar(
        self,
        vmin: float | str | Range | None = None,
        vmax: float | str | Range | None = None,
        label: str | None = None,
        autofit: bool = False,
    ) -> Colorbar:
        row = self.row + 1
        column = self.column + self.shape[1] + 2
        length = self.shape[0]

        if vmin is None:
            vmin = self.range
        if vmax is None:
            vmax = self.range

        cb = Colorbar(row, column, length, sheet=self.sheet)
        cb.set(vmin, vmax, label, autofit)
        return cb


def clean_data(data: DataFrame) -> DataFrame:
    data = data.copy()

    if isinstance(data.columns, MultiIndex):
        data.columns = data.columns.droplevel(list(range(1, data.columns.nlevels)))

    if isinstance(data.index, MultiIndex):
        data.index = data.index.droplevel(list(range(1, data.index.nlevels)))

    data.index.name = None

    return data


def facet(
    row: int,
    column: int,
    data: DataFrame,
    index: str | list[str] | None = None,
    columns: str | list[str] | None = None,
    padding: tuple[int, int] = (1, 1),
) -> NDArray:
    frames = []
    for row_, isub in iterrows(data.index, index, row, padding[0] + 1):
        frames.append([])
        for column_, csub in iterrows(data.columns, columns, column, padding[1] + 1):
            sub = xs(data, isub, csub)
            frame = HeatFrame(row_, column_, sub)
            frames[-1].append(frame)

    return np.array(frames)


def iterrows(
    index: Index,
    levels: str | list[str] | None,
    offset: int = 0,
    padding: int = 0,
) -> Iterator[tuple[int, dict[Hashable, Any]]]:
    if levels is None:
        yield offset, {}
        return

    if isinstance(levels, str):
        levels = [levels]

    if levels:
        values = {level: index.get_level_values(level) for level in levels}
        it = DataFrame(values).drop_duplicates().iterrows()

        for k, (i, s) in enumerate(it):
            if not isinstance(i, int):
                raise NotImplementedError

            yield i + offset + k * padding, s.to_dict()


def xs(
    df: DataFrame,
    index: dict[Hashable, Any] | None,
    columns: dict[Hashable, Any] | None,
) -> DataFrame:
    if index:
        for key, value in index.items():
            df = df.xs(value, level=key, axis=0)  # type: ignore

    if columns:
        for key, value in columns.items():
            df = df.xs(value, level=key, axis=1)  # type: ignore

    return df
