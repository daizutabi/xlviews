from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range as RangeImpl
from xlwings.constants import ChartType

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

        if chart_type is not None:
            self.chart_type = chart_type

        if y is not None:
            self.x = x
            self.y = y

        else:
            self.y = x

    @property
    def label(self) -> str:
        return self.api.Name

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

    def set(
        self,
        style: str | int | None = None,
        color: Color | None = None,
        alpha: float | None = None,
        weight: float | None = None,
        marker: str | int | None = None,
        size: int | None = None,
        line: str | None = None,
        label: str | None = None,
    ) -> Self:
        if label is not None:
            self.label = label

        if self.chart_type == ChartType.xlXYScatter:  # marker
            weight = weight or 1
            return self.marker(style or marker, size, color, alpha, weight)

        weight = weight or 2
        return self.line(style or line or "-", weight, color, alpha, marker, size)

    def marker(
        self,
        style: str | int | None = None,
        size: int | None = None,
        color: Color | None = None,
        alpha: float | None = None,
        weight: float | None = None,
    ) -> Self:
        set_marker(self.api, get_marker_style(style), size)

        if color is not None:
            set_fill(self.api, rgb(color), alpha)
            alpha = alpha / 2 if weight and alpha is not None else alpha
            set_line(self.api, get_line_style(""), rgb(color), alpha, weight)

        return self

    def line(
        self,
        style: str | int | None = None,
        weight: float | None = None,
        color: Color | None = None,
        alpha: float | None = None,
        marker: str | int | None = None,
        size: int | None = None,
    ) -> Self:
        if color is not None:
            color = rgb(color)

        set_line(self.api, get_line_style(style), color, alpha, weight)

        if marker:
            set_marker(self.api, get_marker_style(marker), size)
            if color is not None:
                set_fill(self.api, color, alpha)

        return self


def set_marker(api: Any, style: int | None, size: int | None) -> None:
    if style is not None:
        api.MarkerStyle = style
    if size is not None:
        api.MarkerSize = size


def set_fill(api: Any, color: int | None, alpha: float | None) -> None:
    if color is not None:
        api.Format.Fill.BackColor.RGB = color
        api.Format.Fill.ForeColor.RGB = color
    if alpha is not None:
        api.Format.Fill.Transparency = alpha


def set_line(
    api: Any,
    style: int | None,
    color: int | None,
    alpha: float | None,
    weight: float | None,
) -> None:
    if style is None:
        style = api.Border.LineStyle

    if weight is not None:
        api.Format.Line.Weight = weight
    if color is not None:
        api.Format.Line.ForeColor.RGB = color
    if alpha is not None:
        api.Format.Line.Transparency = alpha

    api.Border.LineStyle = style  # must be set after weight and color
