from __future__ import annotations

from typing import TYPE_CHECKING

from pandas import DataFrame, Index, MultiIndex

from xlviews.core.formula import aggregate
from xlviews.core.range import Range
from xlviews.dataframes.colorbar import Colorbar
from xlviews.style import set_color_scale, set_font
from xlviews.utils import suspend_screen_updates

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator, Sequence
    from typing import Any, Literal, Self

    from pandas import Index
    from xlwings import Sheet

    from xlviews.colors import Color


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

    def style(
        self,
        *,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        color: Color | None = None,
    ) -> Self:
        set_font(self.range, size=size, bold=bold, italic=italic, color=color)
        return self

    def number_format(self, number_format: str) -> Self:
        self.range.impl.number_format = number_format
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

    @classmethod
    def facet(
        cls,
        row: int,
        column: int,
        data: DataFrame,
        index: str | list[str] | None = None,
        columns: str | list[str] | None = None,
        padding: tuple[int, int] = (2, 1),
    ) -> Iterator[tuple[dict[Hashable, Any], Self]]:
        for r, ikey in iterrows(data.index, index, row, padding[0] + 1):
            for c, ckey in iterrows(data.columns, columns, column, padding[1] + 1):
                sub = xs(data, ikey, ckey)
                yield (ikey | ckey), cls(r, c, sub)

    @classmethod
    def pair(
        cls,
        row: int,
        column: int,
        data: DataFrame,
        values: str | list[str] | None = None,
        index: str | list[str] | None = None,
        columns: str | list[str] | None = None,
        padding: tuple[int, int] = (2, 1),
        value_name: str = "value",
        axis: Literal[0, 1] | None = None,
    ) -> Iterator[tuple[dict[Hashable, Any], Self]]:
        if values is None:
            values = data.columns.get_level_values(0).unique().to_list()
        elif isinstance(values, str):
            values = [values]

        if axis is None:
            axis = 1 if columns is None else 0

        nr = row
        nc = column

        for value in values:
            sub = xs(data, None, {0: value})

            for key, frame in cls.facet(row, column, sub, index, columns, padding):
                yield {value_name: value} | key, frame
                if axis == 1:
                    nc = max(nc, frame.column + frame.shape[1] + 1 + padding[1])
                else:
                    nr = max(nr, frame.row + frame.shape[0] + 1 + padding[0])

            column = nc
            row = nr


def clean_data(data: DataFrame) -> DataFrame:
    data = data.copy()

    if isinstance(data.columns, MultiIndex):
        data.columns = data.columns.droplevel(list(range(1, data.columns.nlevels)))

    if isinstance(data.index, MultiIndex):
        data.index = data.index.droplevel(list(range(1, data.index.nlevels)))

    data.index.name = None

    return data


def iterrows(
    index: Index,
    levels: int | str | Sequence[int | str] | None,
    offset: int = 0,
    padding: int = 0,
) -> Iterator[tuple[int, dict[Hashable, Any]]]:
    if levels is None:
        yield offset, {}
        return

    if isinstance(levels, int | str):
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
        df = df.xs(tuple(index.values()), 0, tuple(index.keys()))  # type: ignore

    if columns:
        df = df.xs(tuple(columns.values()), 1, tuple(columns.keys()))  # type: ignore

    return df
