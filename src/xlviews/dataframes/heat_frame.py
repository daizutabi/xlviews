from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
import xlwings
from pandas import DataFrame, Index, MultiIndex

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.core.formula import aggregate
from xlviews.core.range import Range
from xlviews.style import (
    set_alignment,
    set_border,
    set_color_scale,
    set_font,
    set_number_format,
)
from xlviews.utils import suspend_screen_updates

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator
    from typing import Any, Literal, Self

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


class Colorbar:
    start: int
    end: int
    offset: int
    orientation: Literal["vertical", "horizontal"] = "vertical"
    sheet: Sheet
    range: Range

    def __init__(
        self,
        row: int,
        column: int,
        length: int,
        orientation: Literal["vertical", "horizontal"] = "vertical",
        sheet: Sheet | None = None,
    ) -> None:
        self.sheet = sheet or xlwings.sheets.active
        self.orientation = orientation

        if orientation == "vertical":
            self.start = row
            self.end = row + length - 1
            self.offset = column
            self.range = Range((self.start, column), (self.end, column), self.sheet)

        else:
            self.start = column
            self.end = column + length - 1
            self.offset = row
            self.range = Range((row, self.start), (row, self.end), self.sheet)

    def set(
        self,
        vmin: float | str | Range | None = None,
        vmax: float | str | Range | None = None,
        label: str | None = None,
        autofit: bool = False,
    ) -> Self:
        if vmin is not None:
            self.vmin = vmin
        if vmax is not None:
            self.vmax = vmax
        if label is not None:
            self.label = label

        self.draw()

        if autofit:
            self.autofit()

        return self

    @property
    def vmin(self) -> Range:
        i = -1 if self.orientation == "vertical" else 0
        return self.range[i]

    @property
    def vmax(self) -> Range:
        i = 0 if self.orientation == "vertical" else -1
        return self.range[i]

    @vmin.setter
    def vmin(self, value: float | str | Range | list[Range]) -> None:
        if isinstance(value, Range | list):
            func = "min" if len(value) > 1 else None
            value = aggregate(func, value, formula=True)

        self.vmin.value = value

    @vmax.setter
    def vmax(self, value: float | str | Range | list[Range]) -> None:
        if isinstance(value, Range | list):
            func = "max" if len(value) > 1 else None
            value = aggregate(func, value, formula=True)

        self.vmax.value = value

    @property
    def label(self) -> Range:
        offset = (-1, 0) if self.orientation == "vertical" else (0, 1)
        return self.vmax.offset(*offset)

    @label.setter
    def label(self, label: str | None) -> None:
        rng = self.label
        rng.value = label
        set_font(rng, bold=True, size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")

    def draw(self) -> None:
        rng = self.range
        set_color_scale(rng, self.vmin, self.vmax)
        set_font(rng, color=rgb("white"), size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")
        ec = rcParams["heat.border.color"]
        set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)

        vmin = self.vmin.get_address()
        vmax = self.vmax.get_address()

        n = self.end - self.start - 1
        for i in range(n):
            value = f"={vmax}+{i + 1}*({vmin}-{vmax})/{n + 1}"
            if self.orientation == "vertical":
                rng = self.sheet.range(self.start + i + 1, self.offset)
            else:
                rng = self.sheet.range(self.offset, self.start + i + 1)

            rng.value = value
            set_font(rng, size=4)
            set_number_format(rng, "0")

    def apply(self, rng: Range) -> None:
        set_color_scale(rng, self.vmin, self.vmax)

    def autofit(self) -> Self:
        if self.orientation == "vertical":
            start = (self.start - 1, self.offset)
            end = (self.end, self.offset)
        else:
            start = (self.offset, self.start)
            end = (self.offset, self.end + 1)

        self.sheet.range(start, end).autofit()
        return self

    def set_adjacent_column_width(self, width: float, offset: int = 1) -> None:
        """Set the width of the adjacent empty column."""
        if self.orientation == "vertical":
            self.range.offset(0, 1).impl.column_width = width
        else:
            self.range.last_cell.offset(0, 2).impl.column_width = width


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
