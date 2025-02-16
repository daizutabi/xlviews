"""Set styles for Range."""

from __future__ import annotations

import itertools
from typing import TYPE_CHECKING

from xlwings.constants import LineStyle, MarkerStyle, ScaleType

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.core.address import reference
from xlviews.utils import set_font_api

if TYPE_CHECKING:
    from xlwings import Range as RangeImpl
    from xlwings import Sheet

    from xlviews.core.range import Range


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


def get_axis_label(axis) -> str | None:  # noqa: ANN001
    if not axis.HasTitle:
        return None

    return axis.AxisTitle.Text


def set_axis_label(
    axis,  # noqa: ANN001
    label: str | tuple[int, int] | Range | RangeImpl | None = None,
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

    name = name or rcParams["chart.font.name"]
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
    name = name or rcParams["chart.font.name"]
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
