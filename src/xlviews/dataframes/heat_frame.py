from __future__ import annotations

from typing import TYPE_CHECKING

from pandas import Index, MultiIndex

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.core.formula import aggregate
from xlviews.core.range import Range
from xlviews.core.style import (
    set_alignment,
    set_border,
    set_color_scale,
    set_font,
    set_number_format,
)
from xlviews.decorators import suspend_screen_updates
from xlviews.utils import iter_group_locs

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from typing import Self

    from pandas import DataFrame, Index
    from xlwings import Range as RangeImpl
    from xlwings import Sheet


class HeatFrame(SheetFrame):
    index: Index
    columns: Index

    @suspend_screen_updates
    def __init__(
        self,
        row: int,
        column: int,
        data: DataFrame,
        sheet: Sheet | None = None,
    ) -> None:
        data = clean_data(data)

        super().__init__(row, column, data, sheet)

        self.columns = data.columns

        set_heat_frame_style(self)
        self.set_adjacent_column_width(1, offset=-1)
        self.vmin = None
        self.vmax = None
        self.set_colorbar()
        set_style(self)

    def value_range(self) -> Range:
        start = self.row + 1, self.column + 1
        end = start[0] + self.shape[0] - 1, start[1] + self.shape[1] - 1
        return Range(start, end, self.sheet)

    @property
    def vmin(self) -> RangeImpl:
        return self.cell.offset(self.shape[0], self.shape[1] + 2)

    @property
    def vmax(self) -> RangeImpl:
        return self.cell.offset(1, self.shape[1] + 2)

    @property
    def label(self) -> RangeImpl:
        return self.cell.offset(0, self.shape[1] + 2)

    @vmin.setter
    def vmin(self, value: float | str | None) -> None:
        rng = self.value_range()

        if value is None:
            value = aggregate("min", rng, formula=True)

        self.vmin.value = value

    @vmax.setter
    def vmax(self, value: float | str | None) -> None:
        rng = self.value_range()

        if value is None:
            value = aggregate("max", rng, formula=True)

        self.vmax.value = value

    @label.setter
    def label(self, label: str | None) -> None:
        rng = self.label
        rng.value = label
        set_font(rng, bold=True, size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")

    def set_colorbar(self) -> None:
        vmin = self.vmin.get_address()
        vmax = self.vmax.get_address()

        col = self.vmax.column
        start = self.vmax.row
        end = self.vmin.row
        n = end - start - 1
        for i in range(n):
            value = f"={vmax}+{i + 1}*({vmin}-{vmax})/{n + 1}"
            self.sheet.range(i + start + 1, col).value = value

        rng = self.sheet.range((start, col), (end, col))
        set_color_scale(rng, self.vmin, self.vmax)
        set_font(rng, color=rgb("white"), size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")
        ec = rcParams["heat.border.color"]
        set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)

        if n > 0:
            rng = self.sheet.range((start + 1, col), (end - 1, col))
            set_font(rng, size=4)
            set_number_format(rng, "0")

    def autofit(self) -> Self:
        start = self.cell
        end = start.offset(*self.shape)
        self.sheet.range(start, end).autofit()
        self.label.expand("down").autofit()
        return self

    def set_adjacent_column_width(self, width: float, offset: int = 1) -> None:
        """Set the width of the adjacent empty column."""
        column = self.vmax.column + offset
        self.sheet.range(1, column).column_width = width


def clean_data(data: DataFrame) -> DataFrame:
    data = data.copy()

    if isinstance(data.columns, MultiIndex):
        data.columns = data.columns.droplevel(list(range(1, data.columns.nlevels)))

    if isinstance(data.index, MultiIndex):
        data.index = data.index.droplevel(list(range(1, data.index.nlevels)))

    data.index.name = None

    return data


def set_style(sf: HeatFrame) -> None:
    set_color_scale(sf.value_range(), sf.vmin, sf.vmax)
    _merge_index(sf.columns, sf.row, sf.column, 1, sf.sheet)
    _merge_index(sf.index, sf.row, sf.column, 0, sf.sheet)
    _set_border(sf)


def _merge_index(index: Index, row: int, column: int, axis: int, sheet: Sheet) -> None:
    for start, end in iter_group_locs(index):
        if start == end:
            continue
        if axis == 0:
            sheet.range((row + start + 1, column), (row + end + 1, column)).merge()
        else:
            sheet.range((row, column + start + 1), (row, column + end + 1)).merge()


def _set_border(sf: HeatFrame) -> None:
    r = sf.row + 1
    c = sf.column + 1

    ec = rcParams["heat.border.color"]

    for row in iter_group_locs(sf.index, offset=r):
        if row[0] == row[1]:
            continue

        for col in iter_group_locs(sf.columns, offset=c):
            if col[0] == col[1]:
                continue

            rng = sf.sheet.range((row[0], col[0]), (row[1], col[1]))
            set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)
