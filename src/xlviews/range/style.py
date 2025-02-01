"""Set styles for Range."""

from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings as xw
from xlwings import Range, Sheet
from xlwings.constants import (
    BordersIndex,
    ConditionValueTypes,
    FormatConditionType,
    LineStyle,
)

from xlviews.config import rcParams
from xlviews.utils import constant, rgb, set_font_api

if TYPE_CHECKING:
    from .range_collection import RangeCollection


def set_border_line(
    rng: Range,
    index: str,
    weight: int = 2,
    color: int | str = 0,
) -> None:
    if not weight:
        return

    borders = rng.api.Borders
    border = borders(getattr(BordersIndex, index))
    border.LineStyle = LineStyle.xlContinuous
    border.Weight = weight
    border.Color = rgb(color)


def set_border_edge(
    rng: Range,
    weight: int | tuple[int, int, int, int] = 3,
    color: int | str = 0,
) -> None:
    if isinstance(weight, int):
        wl = wr = wt = wb = weight
    else:
        wl, wr, wt, wb = weight

    sheet = rng.sheet
    start, end = rng[0], rng[-1]

    left = sheet.range((start.row, start.column - 1), (end.row, start.column))
    set_border_line(left, "xlInsideVertical", weight=wl, color=color)

    right = sheet.range((start.row, end.column), (end.row, end.column + 1))
    set_border_line(right, "xlInsideVertical", weight=wr, color=color)

    top = sheet.range((start.row - 1, start.column), (start.row, end.column))
    set_border_line(top, "xlInsideHorizontal", weight=wt, color=color)

    bottom = sheet.range((end.row, start.column), (end.row + 1, end.column))
    set_border_line(bottom, "xlInsideHorizontal", weight=wb, color=color)


def set_border_inside(rng: Range, weight: int = 1, color: int | str = 0) -> None:
    set_border_line(rng, "xlInsideVertical", weight=weight, color=color)
    set_border_line(rng, "xlInsideHorizontal", weight=weight, color=color)


def set_border(
    rng: Range,
    edge_weight: int | tuple[int, int, int, int] = 2,
    inside_weight: int = 1,
    edge_color: int | str = 0,
    inside_color: int | str = rgb(140, 140, 140),
) -> None:
    if edge_weight:
        set_border_edge(rng, edge_weight, edge_color)

    if inside_weight:
        set_border_inside(rng, inside_weight, inside_color)


def set_fill(rng: Range | RangeCollection, color: int | str | None = None) -> None:
    if color is not None:
        rng.api.Interior.Color = rgb(color)


def set_font(
    rng: Range | RangeCollection,
    name: str | None = None,
    *,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: int | str | None = None,
) -> None:
    name = name or rcParams["frame.font.name"]
    set_font_api(rng.api, name, size=size, bold=bold, italic=italic, color=color)


def set_alignment(
    rng: Range | RangeCollection,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
) -> None:
    if horizontal_alignment:
        rng.api.HorizontalAlignment = constant(horizontal_alignment)

    if vertical_alignment:
        rng.api.VerticalAlignment = constant(vertical_alignment)


def set_number_format(rng: Range | RangeCollection, fmt: str) -> None:
    if isinstance(rng, Range):
        rng.number_format = fmt
    else:
        for r in rng:
            r.number_format = fmt


def set_banding(
    rng: Range,
    axis: int = 0,
    even_color: int | str = rgb(240, 250, 255),
    odd_color: int | str = rgb(255, 255, 255),
) -> None:
    def banding(mod: int, color: int) -> None:
        formula = f"=MOD(ROW(), 2)={mod}" if axis == 0 else f"=MOD(COLUMN(), 2)={mod}"
        condition = add(Type=FormatConditionType.xlExpression, Formula1=formula)

        condition.SetFirstPriority()
        condition.StopIfTrue = False

        interior = condition.Interior
        interior.PatternColorIndex = constant("automatic")
        interior.Color = color
        interior.TintAndShade = 0

    add = rng.api.FormatConditions.Add

    banding(0, rgb(odd_color))
    banding(1, rgb(even_color))


def hide_succession(rng: Range, color: int | str = rgb(200, 200, 200)) -> None:
    cell = rng[0].get_address(row_absolute=False, column_absolute=False)

    start = rng[0].offset(-2).get_address(column_absolute=False)
    column = rng[0].offset(-1)
    column = ":".join(
        [
            column.get_address(column_absolute=False),
            column.get_address(row_absolute=False, column_absolute=False),
        ],
    )

    ref = (
        f"INDIRECT(ADDRESS(MAX(INDEX(SUBTOTAL(3,OFFSET({start},"
        f'ROW(INDIRECT("1:"&ROWS({column}))),))*ROW({column}),)),'
        f"COLUMN({column})))"
    )
    formula = f"={cell}={ref}"

    add = rng.api.FormatConditions.Add
    condition = add(Type=FormatConditionType.xlExpression, Formula1=formula)
    condition.SetFirstPriority()
    condition.StopIfTrue = False
    condition.Font.Color = rgb(color)


def hide_unique(rng: Range, length: int, color: int | str = rgb(100, 100, 100)) -> None:
    def address(r: Range) -> str:
        return r.get_address(row_absolute=False, column_absolute=False)

    start = rng[0, 0].offset(1, 0)
    end = rng[0, 0].offset(length, 0)
    cell = address(Range(start, end))
    ref = address(start)
    formula = f"=COUNTIF({cell}, {ref}) = {length}"

    add = rng.api.FormatConditions.Add
    condition = add(Type=FormatConditionType.xlExpression, Formula1=formula)
    condition.SetFirstPriority()
    condition.StopIfTrue = False
    condition.Font.Color = rgb(color)
    condition.Font.Italic = True


def hide_gridlines(sheet: Sheet | None = None) -> None:
    sheet = sheet or xw.sheets.active
    sheet.book.app.api.ActiveWindow.DisplayGridlines = False


def set_color_condition(rng: Range, values: list[str], colors: list[int]) -> None:
    condition = rng.api.FormatConditions.AddColorScale(len(values))
    condition.SetFirstPriority()

    for k, (value, color) in enumerate(zip(values, colors, strict=True)):
        criteria = condition.ColorScaleCriteria(k + 1)
        criteria.Type = ConditionValueTypes.xlConditionValueNumber
        criteria.Value = value
        criteria.FormatColor.Color = color


def set_color_scale(
    rng: Range,
    vmin: float | str | Range,
    vmax: float | str | Range,
) -> None:
    if isinstance(vmin, Range):
        vmin = vmin.get_address()

    if isinstance(vmax, Range):
        vmax = vmax.get_address()

    values = [f"={vmin}", f"=({vmin} + {vmax}) / 2", f"={vmax}"]
    colors = [rgb(130, 130, 255), rgb(80, 185, 80), rgb(255, 130, 130)]

    set_color_condition(rng, values, colors)
