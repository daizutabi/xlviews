from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.formula import aggregate
from xlviews.range.range import Range
from xlviews.range.style import set_alignment, set_border, set_color_scale, set_font

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from pandas import DataFrame
    from xlwings import Range as RangeImpl
    from xlwings import Sheet


class HeatFrame(SheetFrame):
    @turn_off_screen_updating
    def __init__(
        self,
        *args,
        data: DataFrame,
        x: str,
        y: str,
        value: str,
        vmin: float | None = None,
        vmax: float | None = None,
        sheet: Sheet | None = None,
        style: bool = True,
        autofit: bool = True,
        font_size: int | None = None,
        **kwargs,
    ) -> None:
        df = data.pivot_table(value, y, x, aggfunc=lambda x: x)
        df.index.name = None

        super().__init__(*args, data=df, index=True, sheet=sheet, style=False)

        if style:
            set_heat_frame_style(self, autofit=autofit, font_size=font_size, **kwargs)

        self.set_adjacent_column_width(1, offset=-1)

        self.set_extrema(vmin, vmax)
        self.set_colorbar()

        set_color_scale(self.heat_range(), self.vmin, self.vmax)

        self.set_label(value)

        if autofit:
            self.label.columns.autofit()

    def heat_range(self) -> Range:
        start = self.row + 1, self.column + 1
        end = start[0] + len(self) - 1, start[1] + len(self.columns) - 1
        return Range(start, end, self.sheet)

    @property
    def vmin(self) -> RangeImpl:
        return self.cell.offset(len(self), len(self.columns) + 1)

    @property
    def vmax(self) -> RangeImpl:
        return self.cell.offset(1, len(self.columns) + 1)

    @property
    def label(self) -> RangeImpl:
        return self.cell.offset(0, len(self.columns) + 1)

    def set_extrema(
        self,
        vmin: float | str | None = None,
        vmax: float | str | None = None,
    ) -> None:
        rng = self.heat_range()

        if vmin is None:
            vmin = aggregate("min", rng, formula=True)

        if vmax is None:
            vmax = aggregate("max", rng, formula=True)

        self.vmin.value = vmin
        self.vmax.value = vmax

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
        ec = rcParams["frame.gray.border.color"]
        set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)

        if n > 0:
            rng = self.sheet.range((start + 1, col), (end - 1, col))
            set_font(rng, size=4)

    def set_label(self, label: str) -> None:
        rng = self.label
        rng.value = label
        set_font(rng, bold=True, size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")

    def set_adjacent_column_width(self, width: float, offset: int = 1) -> None:
        """Set the width of the adjacent empty column."""
        column = self.label.column + offset
        self.sheet.range(1, column).column_width = width
