from __future__ import annotations

from typing import TYPE_CHECKING

from pandas import Index, Series

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.formula import aggregate
from xlviews.range.range import Range
from xlviews.range.style import (
    set_alignment,
    set_border,
    set_color_scale,
    set_font,
    set_number_format,
)

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from collections.abc import Iterator

    from pandas import DataFrame
    from xlwings import Range as RangeImpl
    from xlwings import Sheet


class HeatFrame(SheetFrame):
    x: str | list[str]
    y: str | list[str]
    value: str
    df: DataFrame

    @turn_off_screen_updating
    def __init__(
        self,
        *args,
        data: DataFrame,
        value: str,
        x: str | list[str],
        y: str | list[str],
        vmin: float | None = None,
        vmax: float | None = None,
        sheet: Sheet | None = None,
        style: bool = True,
        autofit: bool = True,
        font_size: int | None = None,
        **kwargs,
    ) -> None:
        df = pivot_table(data, value, y, x)

        self.df = df
        self.value = value
        self.x = x
        self.y = y

        super().__init__(*args, data=df, index=True, sheet=sheet, style=False)

        if style:
            set_heat_frame_style(self, autofit=autofit, font_size=font_size, **kwargs)

        self.set_adjacent_column_width(1, offset=-1)

        self.set_extrema(vmin, vmax)
        self.set_colorbar()

        set_color_scale(self.heat_range(), self.vmin, self.vmax)

        self.set_label(value)

        if autofit:
            self.label.expand("down").autofit()

        if style:
            self.set_heat_style()

    def __len__(self) -> int:
        return len(self.df)

    @property
    def columns(self) -> list:
        return [None, *self.df.columns.to_list()]

    @property
    def index(self) -> Index:
        return self.df.index

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
            set_number_format(rng, "0")

    def set_label(self, label: str) -> None:
        rng = self.label
        rng.value = label
        set_font(rng, bold=True, size=rcParams["frame.font.size"])
        set_alignment(rng, horizontal_alignment="center")

    def set_adjacent_column_width(self, width: float, offset: int = 1) -> None:
        """Set the width of the adjacent empty column."""
        column = self.label.column + offset
        self.sheet.range(1, column).column_width = width

    def set_heat_style(self) -> None:
        if isinstance(self.x, list):
            _merge_index(self.df.columns, self.row, self.column, 1, self.sheet)

        if isinstance(self.y, list):
            _merge_index(self.df.index, self.row, self.column, 0, self.sheet)

        if isinstance(self.x, list) and isinstance(self.y, list):
            _set_border(self)


def pivot_table(
    data: DataFrame,
    value: str,
    y: str | list[str],
    x: str | list[str],
    # aggfunc: Callable = lambda x: x,
) -> DataFrame:
    df = data.pivot_table(value, y, x, aggfunc=lambda x: x)

    if isinstance(y, list):
        df.index = df.index.droplevel(list(range(1, len(y))))

    if isinstance(x, list):
        df.columns = df.columns.droplevel(list(range(1, len(x))))

    df.index.name = None

    return df


def _set_border(sf: HeatFrame) -> None:
    r = sf.row + 1
    c = sf.column + 1

    for row in _iter_merge_ranges(sf.df.index):
        for col in _iter_merge_ranges(sf.df.columns):
            start = (r + row[0], c + col[0])
            end = (r + row[1] - 1, c + col[1] - 1)
            rng = sf.sheet.range(start, end)
            set_border(rng, edge_weight=2, edge_color=rgb("#333333"), inside_weight=0)


def _merge_index(index: Index, row: int, column: int, axis: int, sheet: Sheet) -> None:
    for start, end in _iter_merge_ranges(index):
        if axis == 0:
            sheet.range((row + start + 1, column), (row + end, column)).merge()
        else:
            sheet.range((row, column + start + 1), (row, column + end)).merge()


def _iter_merge_ranges(index: Index) -> Iterator[tuple[int, int]]:
    s = Series(index)
    idx = s[~s.duplicated()].index.to_list()
    return zip(idx, [*idx[1:], len(s)], strict=True)
