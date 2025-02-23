from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range as RangeImpl
from xlwings.constants import LineStyle

from xlviews.colors import Color, rgb
from xlviews.core.range import Range
from xlviews.core.range_collection import RangeCollection

from .style import get_line_style, get_marker_style

if TYPE_CHECKING:
    from typing import Any, Self

    from .axes import Axes


class Series:
    axes: Axes
    api: Any
    # label: str

    def __init__(
        self,
        axes: Axes,
        x: Any,
        y: Any | None = None,
        label: str | None = None,
        chart_type: int | None = None,
    ) -> None:
        self.axes = axes
        self.api = axes.chart.api[1].SeriesCollection().NewSeries()
        self.label = label
        # self.name = self.label

        if chart_type is not None:
            self.chart_type = chart_type

        if y is not None:
            self.x = x
            self.y = y

        else:
            self.y = x

    @property
    def label(self) -> str:
        return str(self.api.Name)

    @label.setter
    def label(self, label: str | None) -> None:
        self.api.Name = label or ""

    @property
    def chart_type(self) -> int:
        return self.api.ChartType  # type: ignore

    @chart_type.setter
    def chart_type(self, chart_type: int) -> None:
        self.api.ChartType = chart_type

    @property
    def x(self) -> tuple:
        return self.api.XValues  # type: ignore

    @x.setter
    def x(self, x: Any) -> None:
        if isinstance(x, Range | RangeImpl | RangeCollection):
            self.api.XValues = x.api
        else:
            self.api.XValues = x

    @property
    def y(self) -> tuple:
        return self.api.Values  # type: ignore

    @y.setter
    def y(self, y: Any) -> None:
        if isinstance(y, Range | RangeImpl | RangeCollection):
            self.api.Values = y.api
        else:
            self.api.Values = y

    def delete(self) -> None:
        self.api.Delete()

    def marker(
        self,
        style: str,
        size: int = 6,
        color: Color | None = None,
        alpha: float = 0,
        weight: float = 1,
    ) -> Self:
        set_marker(self.api, get_marker_style(style), size)
        if color is not None:
            set_fill(self.api, rgb(color), alpha)
            alpha = alpha / 2 if weight else alpha
            set_line(self.api, get_line_style(None), rgb(color), weight, alpha)

        return self

    def line(
        self,
        style: str,
        weight: float = 2,
        color: Color | None = None,
        alpha: float = 0,
        marker: str | None = None,
        size: int = 6,
    ) -> None:
        if color is not None:
            color = rgb(color)

        set_line(self.api, get_line_style(style), color, weight, alpha)

        if marker:
            set_marker(self.api, get_marker_style(marker), size)
            if color is not None:
                set_fill(self.api, color, alpha)


def set_marker(api: Any, style: int, size: int) -> None:
    api.MarkerStyle = style
    api.MarkerSize = size


def set_fill(api: Any, color: int, alpha: float) -> None:
    api.Format.Fill.Visible = True
    api.Format.Fill.BackColor.RGB = color
    api.Format.Fill.Transparency = alpha
    api.Format.Fill.ForeColor.RGB = color


def set_line(
    api: Any,
    style: int,
    color: int | None,
    weight: float,
    alpha: float,
) -> None:
    api.Border.LineStyle = LineStyle.xlContinuous
    api.Format.Line.Visible = True
    api.Format.Line.Weight = weight
    if color is not None:
        api.Format.Line.Transparency = alpha
        api.Format.Line.ForeColor.RGB = color
    api.Border.LineStyle = style
