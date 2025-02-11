from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.formula import Func, aggregate
from xlviews.range.range import Range
from xlviews.range.style import (
    set_alignment,
    set_border,
    set_color_scale,
    set_font,
    set_number_format,
)
from xlviews.utils import iter_group_ranges

from .sheet_frame import SheetFrame
from .style import set_heat_frame_style

if TYPE_CHECKING:
    from typing import Self

    from pandas import DataFrame, Index
    from xlwings import Range as RangeImpl
    from xlwings import Sheet


class HeatFrame(SheetFrame):
    _data: DataFrame

    @turn_off_screen_updating
    def __init__(
        self,
        row: int,
        column: int,
        data: DataFrame,
        vmin: float | None = None,
        vmax: float | None = None,
        sheet: Sheet | None = None,
    ) -> None:
        # sheet = sheet or (
        #     data.sheet if isinstance(data, SheetFrame) else xlwings.sheets.active
        # )

        # if isinstance(data, SheetFrame):
        #     data = _to_dataframe(
        #         data,
        #         value,
        #         _to_list(x, y),
        #         aggfunc,
        #         include_sheetname=sheet.name != data.sheet.name,
        #     )

        # if value and x and y:
        #     df = pivot_table(data, value, x, y)
        # else:
        #     df = data

        self._data = data

        super().__init__(row, column, data=data, index=True, sheet=sheet)

        set_heat_frame_style(self)
        self.set_adjacent_column_width(1, offset=-1)
        self.set_extrema(vmin, vmax)
        self.set_colorbar()
        set_color_scale(self.heat_range(), self.vmin, self.vmax)
        self.set_heat_style()

        # self.set_label(value)

        # if autofit:
        #     self.label.expand("down").autofit()

        # if style:
        #     self.set_heat_style()

    @property
    def shape(self) -> tuple[int, int]:
        return self._data.shape

    def __len__(self) -> int:
        return self.shape[0]

    # @property
    # def index(self) -> Index:
    #     return self.data.index

    def heat_range(self) -> Range:
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

    def set_label(self, label: str | None) -> None:
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

    def set_heat_style(self) -> None:
        _merge_index(self.data.columns, self.row, self.column, 1, self.sheet)
        _merge_index(self.data.index, self.row, self.column, 0, self.sheet)
        _set_border(self)

    @classmethod
    def facet(
        cls,
        cell_row: int,
        cell_column: int,
        data: DataFrame | SheetFrame,
        value: str,
        x: str | list[str] | None,
        y: str | list[str] | None,
        col: str | list[str],
        row: str | list[str],
        aggfunc: Func = None,
        vmin: float | None = None,
        vmax: float | None = None,
        sheet: Sheet | None = None,
        style: bool = True,
        autofit: bool = True,
        font_size: int | None = None,
        **kwargs,
    ) -> list[Self]:
        sheet = sheet or (
            data.sheet if isinstance(data, SheetFrame) else xlwings.sheets.active
        )

        if isinstance(data, SheetFrame):
            data = _to_dataframe(
                data,
                value,
                _to_list(col, row, x, y),
                aggfunc,
                include_sheetname=sheet.name != data.sheet.name,
            )

        df = pivot_table(
            data,
            value,
            _to_list(col, x),
            _to_list(row, y),
            drop_levels=False,
        )
        print(df)
        print(df.loc[[1], [1]])

        # cls(
        #     cell_row,
        #     cell_column,
        #     value=None,
        #     x=None,
        #     y=None,
        #     sheet=sheet,
        #     style=style,
        #     autofit=autofit,
        #     font_size=font_size,
        #     **kwargs,
        # )

        # sheet = sheet or (
        #     data.sheet if isinstance(data, SheetFrame) else xlwings.sheets.active
        # )

        # if isinstance(data, SheetFrame):
        #     data = _to_dataframe(
        #         data,
        #         value,
        #         _to_list(x, y),
        #         aggfunc,
        #         include_sheetname=sheet.name != data.sheet.name,
        #     )

        # if value and x and y:
        #     df = pivot_table(data, value, x, y)
        # else:
        #     df = data

        # self.df = df


def _to_list(*args: str | list[str] | None) -> list[str]:
    results = []

    for arg in args:
        if arg is None:
            continue

        if isinstance(arg, list):
            results.extend(arg)
        else:
            results.append(arg)

    return results


def _to_dataframe(
    sf: SheetFrame,
    value: str | None,
    by: str | list[str] | None,
    aggfunc: Func = None,
    include_sheetname: bool = False,
) -> DataFrame:
    columns = [value] if isinstance(value, str) else value

    if aggfunc is None:
        return sf.get_address(
            columns,
            include_sheetname=include_sheetname,
            formula=True,
        )

    return sf.groupby(by).agg(
        aggfunc,
        columns,
        include_sheetname=include_sheetname,
        formula=True,
    )


def pivot_table(
    data: DataFrame,
    value: str,
    x: str | list[str],
    y: str | list[str],
    drop_levels: bool = True,
    # aggfunc: Callable = lambda x: x,
) -> DataFrame:
    df = data.pivot_table(value, y, x, aggfunc=lambda x: x)

    if isinstance(x, list) and drop_levels:
        df.columns = df.columns.droplevel(list(range(1, len(x))))

    if isinstance(y, list) and drop_levels:
        df.index = df.index.droplevel(list(range(1, len(y))))

    df.index.name = None

    return df


def _set_border(sf: HeatFrame) -> None:
    r = sf.row + 1
    c = sf.column + 1

    ec = rcParams["heat.border.color"]

    for row in iter_group_ranges(sf.data.index):
        if row[0] == row[1]:
            continue

        for col in iter_group_ranges(sf.data.columns):
            if col[0] == col[1]:
                continue

            start = (r + row[0], c + col[0])
            end = (r + row[1], c + col[1])
            rng = sf.sheet.range(start, end)
            set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)


def _merge_index(index: Index, row: int, column: int, axis: int, sheet: Sheet) -> None:
    for start, end in iter_group_ranges(index):
        if start == end:
            continue
        if axis == 0:
            sheet.range((row + start + 1, column), (row + end + 1, column)).merge()
        else:
            sheet.range((row, column + start + 1), (row, column + end + 1)).merge()
