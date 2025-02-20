from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings

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

if TYPE_CHECKING:
    from typing import Literal, Self

    from xlwings import Sheet


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
        vmin: float | str | Range | list[Range] | None = None,
        vmax: float | str | Range | list[Range] | None = None,
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
