"""Set styles such as Marker."""

from __future__ import annotations

import itertools
from functools import partial
from typing import TYPE_CHECKING

import pywintypes
import seaborn as sns
from xlwings import Range, Sheet
from xlwings.constants import (
    BordersIndex,
    FormatConditionType,
    LineStyle,
    MarkerStyle,
    ScaleType,
    TableStyleElementType,
)

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range import reference
from xlviews.utils import constant, rgb

if TYPE_CHECKING:
    from xlwings._xlwindows import COMRetryObjectWrapper

    from xlviews.sheetframe import SheetFrame
    from xlviews.table import Table


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


def set_fill(rng: Range, color: int | str | None = None) -> None:
    if color is not None:
        rng.api.Interior.Color = rgb(color)


def set_font_api(
    api: COMRetryObjectWrapper,
    name: str | None = None,
    *,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: int | str | None = None,
) -> None:
    name = name or rcParams["chart.font.name"]

    font = api.Font
    font.Name = name  # type: ignore
    if size:
        font.Size = size  # type: ignore
    if bold is not None:
        font.Bold = bold  # type: ignore
    if italic is not None:
        font.Italic = italic  # type: ignore
    if color is not None:
        font.Color = rgb(color)  # type: ignore


def set_font(rng: Range, *args, **kwargs) -> None:
    set_font_api(rng.api, *args, **kwargs)


def set_alignment(
    rng: Range,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
) -> None:
    if horizontal_alignment:
        rng.api.HorizontalAlignment = constant(horizontal_alignment)

    if vertical_alignment:
        rng.api.VerticalAlignment = constant(vertical_alignment)


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


def hide_gridlines(sheet: Sheet) -> None:
    sheet.book.app.api.ActiveWindow.DisplayGridlines = False


def _set_style(
    start: Range,
    end: Range,
    name: str,
    *,
    border: bool = True,
    gray: bool = False,
    font: bool = True,
    fill: bool = True,
    font_size: int | None = None,
) -> None:
    rng = start.sheet.range(start, end)

    if border:
        set_border(rng, edge_color=rcParams["frame.gray.border.color"] if gray else 0)

    if fill:
        _set_style_fill(rng, name, gray=gray)

    if font:
        _set_style_font(rng, name, gray=gray, font_size=font_size)


def _set_style_fill(rng: Range, name: str, *, gray: bool = False) -> None:
    if gray and name != "values":
        color = rcParams["frame.gray.fill.color"]
    else:
        color = rcParams[f"frame.{name}.fill.color"]

    set_fill(rng, color=color)


def _set_style_font(
    rng: Range,
    name: str,
    *,
    gray: bool = False,
    font_size: int | None = None,
) -> None:
    if gray:
        color = rcParams["frame.gray.font.color"]
    else:
        color = rcParams[f"frame.{name}.font.color"]
    bold = rcParams[f"frame.{name}.font.bold"]
    size = font_size or rcParams["frame.font.size"]

    set_font(rng, color=color, bold=bold, size=size)


@turn_off_screen_updating
def set_frame_style(
    sf: SheetFrame,
    *,
    autofit: bool = False,
    alignment: str | None = "center",
    banding: bool = False,
    succession: bool = False,
    border: bool = True,
    gray: bool = False,
    font: bool = True,
    fill: bool = True,
    font_size: int | None = None,
) -> None:
    """Set style of SheetFrame.

    Args:
        sf: The SheetFrame object.
        autofit: Whether to autofit the frame.
        alignment: The alignment of the frame.
        border: Whether to draw the border.
        font: Whether to specify the font.
        fill: Whether to fill the frame.
        banding: Whether to draw the banding.
        succession: Whether to hide the succession of the index.
        gray: Whether to set the frame in gray mode.
        font_size: The font size to specify directly.
    """
    cell = sf.cell
    sheet = sf.sheet

    set_style = partial(
        _set_style,
        border=border,
        gray=gray,
        font=font,
        fill=fill,
        font_size=font_size,
    )

    index_level = sf.index_level
    columns_level = sf.columns_level
    length = len(sf)

    if index_level > 0:
        start = cell
        end = cell.offset(columns_level - 1, index_level - 1)
        set_style(start, end, "index.name")

        start = cell.offset(columns_level, 0)
        end = cell.offset(columns_level + length - 1, index_level - 1)
        set_style(start, end, "index")

        if succession:
            rng = sheet.range(start.offset(1, 0), end)
            hide_succession(rng)

            start = cell.offset(columns_level - 1, 0)
            end = cell.offset(columns_level - 1, index_level - 1)
            rng = sheet.range(start, end)
            hide_unique(rng, length)

    width = len(sf.value_columns)

    if columns_level > 1:
        start = cell.offset(0, index_level)
        end = cell.offset(columns_level - 2, index_level + width - 1)
        set_style(start, end, "columns.name")

    start = cell.offset(columns_level - 1, index_level)
    end = cell.offset(columns_level - 1, index_level + width - 1)
    set_style(start, end, "columns")

    start = cell.offset(columns_level, index_level)
    end = cell.offset(columns_level + length - 1, index_level + width - 1)
    set_style(start, end, "values")

    rng = sheet.range(start, end)

    if banding and not gray:
        set_banding(rng)

    rng = sheet.range(cell, end)

    if border:
        ew = 2 if gray else 3
        ec = rcParams["frame.gray.border.color"] if gray else 0
        set_border(rng, edge_weight=ew, inside_weight=0, edge_color=ec)

    if autofit:
        rng.columns.autofit()

    if alignment:
        set_alignment(rng, alignment)


def set_wide_column_style(sf: SheetFrame, gray: bool = False) -> None:
    wide_columns = sf.wide_columns
    edge_color = rcParams["frame.gray.border.color"] if gray else 0

    for wide_column in wide_columns:
        rng = sf.range(wide_column).offset(-1)

        er = 3 if wide_column == wide_columns[-1] else 2
        edge_weight = (1, er - 1, 1, 1) if gray else (2, er, 2, 2)
        set_border(rng, edge_weight, inside_weight=1, edge_color=edge_color)

        _set_style_fill(rng, "wide-columns", gray=gray)
        _set_style_font(rng, "wide-columns", gray=gray)

    for wide_column in wide_columns:
        rng = sf.range(wide_column).offset(-2)

        el = 3 if wide_column == wide_columns[0] else 2
        edge_weight = (el - 1, 2, 2, 1) if gray else (el, 3, 3, 2)
        set_border(rng, edge_weight, inside_weight=0, edge_color=edge_color)

        _set_style_fill(rng, "wide-columns.name", gray=gray)
        _set_style_font(rng, "wide-columns.name", gray=gray)


def set_table_style(
    table: Table,
    even_color: int | str = rgb(240, 250, 255),
    odd_color: int | str = rgb(255, 255, 255),
) -> None:
    book = table.sheet.book.api

    try:
        style = book.TableStyles("xlviews")
    except pywintypes.com_error:
        style = book.TableStyles.Add("xlviews")
        odd_type = TableStyleElementType.xlRowStripe1
        style.TableStyleElements(odd_type).Interior.Color = odd_color
        even_type = TableStyleElementType.xlRowStripe2
        style.TableStyleElements(even_type).Interior.Color = even_color

    table.api.TableStyle = style


def color_palette(n: int) -> list[tuple[int, int, int]]:
    """Return a list of colors of length n."""
    palette = sns.color_palette()
    palette = palette[:n] if n <= len(palette) else sns.husl_palette(n, l=0.5)
    return [tuple(int(c * 255) for c in p) for p in palette]  # type: ignore


MARKER_DICT: dict[str, int] = {
    "o": MarkerStyle.xlMarkerStyleCircle,
    "^": MarkerStyle.xlMarkerStyleTriangle,
    "s": MarkerStyle.xlMarkerStyleSquare,
    "d": MarkerStyle.xlMarkerStyleDiamond,
    "+": MarkerStyle.xlMarkerStylePlus,
    "x": MarkerStyle.xlMarkerStyleX,
    ".": MarkerStyle.xlMarkerStyleDot,
    "-": MarkerStyle.xlMarkerStyleDash,
    "*": MarkerStyle.xlMarkerStyleStar,
}

LINE_DICT: dict[str, int] = {
    "-": LineStyle.xlContinuous,
    "--": LineStyle.xlDash,
    "-.": LineStyle.xlDashDot,
    ".": LineStyle.xlDot,
}


def get_marker_style(marker: int | str | None) -> int:
    if isinstance(marker, int):
        return marker

    if marker is None:
        return MarkerStyle.xlMarkerStyleNone

    return MARKER_DICT[marker]


def get_line_style(line: int | str | None) -> int:
    if isinstance(line, int):
        return line

    if line is None:
        return LineStyle.xlLineStyleNone

    return LINE_DICT[line]


def marker_palette(n: int) -> list[str]:
    """Return a list of markers of length n."""
    return list(itertools.islice(itertools.cycle(MARKER_DICT), n))


def palette(name: str, n: int) -> list[str] | list[tuple[int, int, int]] | list[None]:
    if name == "color":
        return color_palette(n)

    if name == "marker":
        return marker_palette(n)

    return [None] * n


def get_axis_label(axis) -> str | None:  # noqa: ANN001
    if not axis.HasTitle:
        return None

    return axis.AxisTitle.Text


def set_axis_label(
    axis,  # noqa: ANN001
    label: str | tuple[int, int] | Range | None = None,
    name: str | None = None,
    size: float | None = None,
    sheet: Sheet | None = None,
    **kwargs,
) -> None:
    if not label:
        axis.HasTitle = False
        return

    axis.HasTitle = True
    axis_title = axis.AxisTitle
    axis_title.Text = reference(label, sheet)
    size = size or rcParams["chart.axis.title.font.size"]

    set_font_api(axis_title, name, size=size, **kwargs)


def get_ticks(axis) -> tuple[float, float, float, float]:  # noqa: ANN001
    return (
        axis.MinimumScale,
        axis.MaximumScale,
        axis.MajorUnit,
        axis.MinorUnit,
    )


def set_ticks(
    axis,  # noqa: ANN001
    *args,
    min: float | None = None,  # noqa: A002
    max: float | None = None,  # noqa: A002
    major: float | None = None,
    minor: float | None = None,
    gridlines: bool = True,
) -> None:
    args = [*args, None, None, None, None][:4]

    min = min or args[0]  # noqa: A001
    max = max or args[1]  # noqa: A001
    major = major or args[2]
    minor = minor or args[3]

    if min is not None:
        axis.MinimumScale = min

    if max is not None:
        axis.MaximumScale = max

    if major is not None:
        axis.MajorUnit = major

        if gridlines:
            axis.HasMajorGridlines = True
        else:
            axis.HasMajorGridlines = False

    if minor is not None:
        axis.MinorUnit = minor

        if gridlines:
            axis.HasMinorGridlines = True
        else:
            axis.HasMinorGridlines = False

    if min:
        axis.CrossesAt = min


def set_tick_labels(
    axis,  # noqa: ANN001
    name: str | None = None,
    size: float | None = None,
    number_format: str | None = None,
) -> None:
    size = size or rcParams["chart.axis.ticklabels.font.size"]
    set_font_api(axis.TickLabels, name, size=size)

    if number_format:
        axis.TickLabels.NumberFormatLocal = number_format


def get_axis_scale(axis) -> str:  # noqa: ANN001
    if axis.ScaleType == ScaleType.xlScaleLogarithmic:
        return "log"

    if axis.ScaleType == ScaleType.xlScaleLinear:
        return "linear"

    raise NotImplementedError


def set_axis_scale(axis, scale: str) -> None:  # noqa: ANN001
    if scale == "log":
        axis.ScaleType = ScaleType.xlScaleLogarithmic
        return

    if scale == "linear":
        axis.ScaleType = ScaleType.xlScaleLinear
        return

    raise NotImplementedError


def set_dimensions(
    api,  # noqa: ANN001
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> None:
    if left is not None:
        api.Left = left

    if top is not None:
        api.Top = top

    if width is not None:
        api.Width = width

    if height is not None:
        api.Height = height


def set_area_format(
    api,  # noqa: ANN001
    border: str | int | tuple[int, int, int] | None = None,
    fill: str | int | tuple[int, int, int] | None = None,
    alpha: float | None = None,
) -> None:
    if border is not None:
        api.Format.Line.Visible = True
        api.Format.Line.ForeColor.RGB = rgb(border)

    if fill is not None:
        api.Format.Fill.Visible = True
        api.Format.Fill.ForeColor.RGB = rgb(fill)

    if alpha is not None:
        api.Format.Line.Transparency = alpha
        api.Format.Fill.Transparency = alpha
