from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range as RangeImpl
from xlwings.constants import LineStyle

from xlviews.colors import rgb
from xlviews.core.address import reference
from xlviews.core.range import Range
from xlviews.core.range_collection import RangeCollection

from .style import get_line_style, get_marker_style

if TYPE_CHECKING:
    from typing import Any

    from xlwings import Sheet

    from .axes import Axes


class Series:
    axes: Axes
    api: Any
    label: str

    def __init__(
        self,
        axes: Axes,
        x: Any,
        y: Any | None = None,
        label: str | tuple[int, int] | Range | RangeImpl = "",
        chart_type: int | None = None,
        sheet: Sheet | None = None,
    ) -> None:
        self.axes = axes
        self.api = axes.chart.api[1].SeriesCollection().NewSeries()
        self.label = label if isinstance(label, str) else reference(label, sheet)
        self.name = self.label

        if chart_type is not None:
            self.chart_type = chart_type

        if y is not None:
            self.x = x
            self.y = y

        else:
            self.y = x

    @property
    def name(self) -> str:
        return str(self.api.Name)

    @name.setter
    def name(self, name: str) -> None:
        self.api.Name = name

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
        marker: str | None = "o",
        line: str | None = "-",
        color: str | int | tuple[int, int, int] = "black",
        size: int = 5,
        weight: float = 2,
        alpha: float = 0,
        **kwargs,
    ) -> None:
        if line is None:
            weight = min(size / 4, weight)
            line_alpha = alpha / 2
        else:
            line_alpha = alpha

        set_marker(self.api, get_marker_style(marker), size)
        set_fill(self.api, rgb(color), alpha)
        set_line(self.api, rgb(color), get_line_style(line), weight, line_alpha)


def set_marker(api: Any, style: int, size: int) -> None:
    api.MarkerStyle = style
    api.MarkerSize = size


def set_fill(api: Any, color: int, alpha: float) -> None:
    api.Format.Fill.Visible = True
    api.Format.Fill.BackColor.RGB = color
    api.Format.Fill.Transparency = alpha
    api.Format.Fill.ForeColor.RGB = color


def set_line(api: Any, color: int, style: int, weight: float, alpha: float) -> None:
    api.Border.LineStyle = LineStyle.xlContinuous
    api.Format.Line.Visible = True
    api.Format.Line.Weight = weight
    api.Format.Line.Transparency = alpha
    api.Format.Line.ForeColor.RGB = color
    api.Border.LineStyle = style
