"""Set styles such as Marker."""

from __future__ import annotations

import itertools
from functools import partial
from typing import TYPE_CHECKING

import pywintypes
import seaborn as sns
import xlwings as xw
from xlwings import Range, Sheet
from xlwings.constants import BordersIndex, FormatConditionType, LineStyle

from xlviews.config import rcParams
from xlviews.decorators import api, wait_updating
from xlviews.utils import constant, rgb

if TYPE_CHECKING:
    from xlwings._xlwindows import COMRetryObjectWrapper


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
    size: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: int | str | None = None,
) -> None:
    if name is None:
        name = rcParams["chart.font.name"]

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
        set_border(rng, edge_color="#aaaaaa" if gray else 0)

    if fill:
        if gray and name != "values":
            color = "#eeeeee"
        else:
            color = rcParams[f"frame.{name}.fill.color"]
        set_fill(rng, color=color)

    if font:
        color = "#aaaaaa" if gray else rcParams[f"frame.{name}.font.color"]
        bold = rcParams[f"frame.{name}.font.bold"]
        size = font_size or rcParams["frame.font.size"]
        set_font(rng, color=color, bold=bold, size=size)


@wait_updating
def set_frame_style(
    cell: Range,
    index_level: int,
    columns_level: int,
    length: int,
    columns: int,
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
        cell: The top-left cell of the frame.
        index_level: The depth of the index level.
        columns_level: The depth of the columns level.
        length: The number of rows.
        columns: The number of columns excluding the index.
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
    sheet = cell.sheet
    set_style = partial(
        _set_style,
        border=border,
        gray=gray,
        font=font,
        fill=fill,
        font_size=font_size,
    )

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

    if columns_level > 1:
        start = cell.offset(0, index_level)
        end = cell.offset(columns_level - 2, index_level + columns - 1)
        set_style(start, end, "columns.name")

    start = cell.offset(columns_level - 1, index_level)
    end = cell.offset(columns_level - 1, index_level + columns - 1)
    set_style(start, end, "columns")

    start = cell.offset(columns_level, index_level)
    end = cell.offset(columns_level + length - 1, index_level + columns - 1)
    set_style(start, end, "values")

    rng = sheet.range(start, end)
    if border:
        set_border(rng, edge_color="#aaaaaa" if gray else 0)
    if banding and not gray:
        set_banding(rng)
    if gray:
        set_font(rng, color="#aaaaaa")

    rng = sheet.range(cell, end)
    if border:
        ew = 2 if gray else 3
        ec = "#aaaaaa" if gray else 0
        set_border(rng, edge_weight=ew, inside_weight=0, edge_color=ec)
    if autofit:
        rng.columns.autofit()
    if alignment:
        set_alignment(rng, alignment)


def set_table_style(api, even_color=rgb(240, 250, 255), odd_color=rgb(255, 255, 255)):
    book = api.Range.Parent.Parent
    try:
        style = book.TableStyles("xlviews")
    except pywintypes.com_error:
        style = book.TableStyles.Add("xlviews")
        odd_type = xw.constants.TableStyleElementType.xlRowStripe1
        style.TableStyleElements(odd_type).Interior.Color = odd_color
        even_type = xw.constants.TableStyleElementType.xlRowStripe2
        style.TableStyleElements(even_type).Interior.Color = even_color
    api.TableStyle = style


def color_palette(n: int) -> list[tuple[int, int, int]]:
    """Return a list of colors of length n."""
    palette = sns.color_palette()
    palette = palette[:n] if n <= len(palette) else sns.husl_palette(n, l=0.5)
    return [tuple(int(c * 255) for c in p) for p in palette]  # type: ignore


MARKER_DICT: dict[str, str] = {
    "o": "circle",
    "^": "triangle",
    "s": "square",
    "d": "diamond",
    "+": "plus",
    "x": "x",
    ".": "dot",
    "-": "dash",
    "*": "star",
}

LINE_DICT: dict[str, str] = {
    "-": "continuous",
    "--": "dash",
    "-.": "dashDot",
    ".": "Dot",
}


def marker_palette(n: int) -> list[str]:
    """Return a list of markers of length n."""
    return list(itertools.islice(itertools.cycle(MARKER_DICT), n))


def palette(name: str, n: int) -> list[str] | list[tuple[int, int, int]] | list[None]:
    if name == "color":
        return color_palette(n)

    if name == "marker":
        return marker_palette(n)

    return [None] * n


@api
def set_series_style(
    series,
    marker=False,
    size=False,
    line=False,
    color=False,
    fill_color=False,
    edge_color=False,
    line_color=False,
    width=False,
    edge_width=False,
    line_width=False,
    alpha=False,
    fill_alpha=False,
    edge_alpha=False,
    line_alpha=False,
):
    """
    Seriesのスタイルを設定する.
    Noneが有効な指定であるため、指定しないことを示すデフォルト値をFalseとする。
    """
    # size = 10
    # edge_width = 3
    fill = series.Format.Fill
    edge = series.Format.Line
    border = series.Border

    has_line = line or border.LineStyle != xw.constants.LineStyle.xlLineStyleNone
    has_marker = (
        marker or series.MarkerStyle != xw.constants.MarkerStyle.xlMarkerStyleNone
    )

    # 'is not False' は 0 が有効な指定であるため
    if color is not False and color is not None:
        if line_color is False and has_line:
            line_color = color
        if fill_color is False and has_marker:
            fill_color = color
        if edge_color is False and has_marker:
            edge_color = color

    if alpha is not False and alpha is not None:
        if line_alpha is False and has_line:
            line_alpha = alpha
        if fill_alpha is False and has_marker:
            fill_alpha = alpha
        if edge_alpha is False and has_marker:
            edge_alpha = alpha / 2

    if marker is None:
        series.MarkerStyle = xw.constants.MarkerStyle.xlMarkerStyleNone
    elif marker:
        marker = MARKER_DICT.get(marker, marker)
        marker = "xlMarkerStyle" + marker[0].upper() + marker[1:]
        marker = getattr(xw.constants.MarkerStyle, marker)
        series.MarkerStyle = marker
    if size:
        series.MarkerSize = size

    # 以下の通りの順番に実行することが重要！！
    # edge を指定すると、lineの変わってしまうため覚えておく
    line_style = border.LineStyle

    if fill_color is not False:
        fill.Visible = True
        fill.BackColor.RGB = rgb(fill_color)
    if fill_alpha is not False:
        fill.Transparency = fill_alpha
    if fill_color is not False:
        fill.ForeColor.RGB = rgb(fill_color)

    if edge_color is not False:
        edge.Visible = True
        edge.BackColor.RGB = rgb(edge_color)
    if edge_alpha is not False:
        edge.Transparency = edge_alpha
        # lineとedgeの透明度は独立に指定する方法が分からない。そのため、
        # lineの透明度を指定したときにはマーカーのエッジを消す。
        line_width_ = border.Weight
        edge.Weight = 0
        border.Weight = line_width_
    if edge_color is not False:
        edge.ForeColor.RGB = rgb(edge_color)
    if edge_width is not False:
        edge.Weight = edge_width

    if line is False:
        border.LineStyle = line_style
    elif line is None:
        border.LineStyle = xw.constants.LineStyle.xlLineStyleNone
    elif line:
        line = LINE_DICT.get(line, line)
        line = "xl" + line[0].upper() + line[1:]
        line = getattr(xw.constants.LineStyle, line)
        border.LineStyle = line

    if line_color is not False:
        border.Color = rgb(line_color)
    if line_alpha is not False:
        edge.Transparency = line_alpha
        # lineとedgeの透明度は独立に指定する方法が分からない。そのため、
        # lineの透明度を指定したときにはマーカーのエッジを消す。
        line_width_ = border.Weight
        edge.Weight = 0
        border.Weight = line_width_
    if line_width is not False:
        border.Weight = line_width

    if line is None:
        edge.Visible = False


@api
def set_scale(axis, scale):
    if not scale:
        return
    if scale == "log":
        axis.ScaleType = xw.constants.ScaleType.xlScaleLogarithmic
    elif scale == "linear":
        axis.ScaleType = xw.constants.ScaleType.xlScaleLinear


@api
def set_label(axis, label, size=None, name=None, **kwargs):
    if not label:
        axis.HasTitle = False
        return
    axis.HasTitle = True
    axis_title = axis.AxisTitle
    axis_title.Text = label
    if size is None:
        size = rcParams["chart.axis.title.font.size"]
    set_font(axis_title, size=size, name=name, **kwargs)


@api
def set_ticks(
    axis,
    *args,
    min=None,
    max=None,
    major=None,
    minor=None,
    gridlines=True,
    **kwargs,
):
    args = (list(args) + [None, None, None, None])[:4]
    min = min or args[0]
    max = max or args[1]
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


@api
def set_ticklabels(axis, name=None, size=None, format=None):
    if size is None:
        size = rcParams["chart.axis.ticklabels.font.size"]
    set_font(axis.TickLabels, name=name, size=size)
    # set_font(axis.Format.TextFrame2.TextRange, name=name, size=size)
    if format:
        axis.TickLabels.NumberFormatLocal = format


def set_dimensions(obj, left=None, top=None, width=None, height=None):
    if left is not None:
        obj.Left = left
    if top is not None:
        obj.Top = top
    if width is not None:
        obj.Width = width
    if height is not None:
        obj.Height = height


def set_area(obj, border=None, fill=None, alpha=None):
    if border is not None:
        obj.Format.Line.Visible = True
        obj.Format.Line.ForeColor.RGB = rgb(border)
    if fill is not None:
        obj.Format.Fill.Visible = True
        obj.Format.Fill.ForeColor.RGB = rgb(fill)
    if alpha is not None:
        obj.Format.Line.Transparency = alpha
        obj.Format.Fill.Transparency = alpha
