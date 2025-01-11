from __future__ import annotations

from contextlib import suppress
from typing import TYPE_CHECKING

import xlwings as xw
from xlwings import Range
from xlwings.constants import AxisType, Placement, TickMark

from xlviews.config import rcParams
from xlviews.range import reference
from xlviews.series import Series
from xlviews.style import (
    get_axis_label,
    get_axis_scale,
    get_ticks,
    set_area_format,
    set_axis_label,
    set_axis_scale,
    set_dimensions,
    set_font_api,
    set_tick_labels,
    set_ticks,
)

if TYPE_CHECKING:
    from typing import Any

    from xlwings import Chart, Sheet

    from xlviews.sheetframe import SheetFrame

FIRST_POSITION = {"left": 50, "top": 50}


def clear_first_position(left: int = 50, top: int = 50) -> None:
    FIRST_POSITION["left"] = left
    FIRST_POSITION["top"] = top


def set_first_position(sf: SheetFrame, pos: str = "right") -> None:
    if pos == "right":
        cell = sf.get_adjacent_cell(offset=0)
        FIRST_POSITION["left"] = cell.left
        FIRST_POSITION["top"] = cell.top

    elif pos == "inside":
        cell = sf.cell.offset(sf.columns_level, sf.index_level)
        FIRST_POSITION["left"] = cell.left + 30
        FIRST_POSITION["top"] = cell.top + 30

    elif pos == "bottom":
        cell = sf.cell.offset(sf.columns_level + len(sf) + 1)
        FIRST_POSITION["left"] = cell.left
        FIRST_POSITION["top"] = cell.top


def chart_position(
    sheet: Sheet,
    left: float | None,
    top: float | None,
) -> tuple[float, float]:
    """Return the position of the chart.

    If left is 0 and top is None, it will create a new row.
    If left is None and top is None, it will be placed to the right.
    """
    if left is not None and top is not None:
        return left, top

    if not sheet.charts:
        return left or FIRST_POSITION["left"], top or FIRST_POSITION["top"]

    if left == 0 and top is None:  # New row
        left = FIRST_POSITION["left"]
        top = FIRST_POSITION["top"]

        for chart in sheet.charts:
            top = max(top, chart.top)
            left = chart.left if left < 0 else min(left, chart.left)

        for chart in sheet.charts:
            if chart.top == top:
                top = max(top, chart.top + chart.height)

        return left, top

    chart = sheet.charts[-1]
    return chart.left + chart.width, chart.top


class Axes:
    sheet: Sheet
    chart: Chart
    chart_type: int
    series_collection: list[Series]

    def __init__(
        self,
        left: float | None = None,
        top: float | None = None,
        width: float = 0,
        height: float = 0,
        *,
        row: int | None = None,
        column: int | None = None,
        sheet: Sheet | None = None,
        chart_type: int | None = None,
        border_width: int = 0,
        visible_only: bool = True,
        has_legend: bool = True,
        include_in_layout: bool = False,
    ) -> None:
        self.sheet = sheet or xw.sheets.active

        if row:
            top = self.sheet.range(row, 1).top
        if column:
            left = self.sheet.range(1, column).left

        left, top = chart_position(self.sheet, left, top)

        width = width or rcParams["chart.width"]
        height = height or rcParams["chart.height"]

        self.chart = self.sheet.charts.add(left, top, width, height)  # type: ignore

        if chart_type is None:
            self.chart_type = self.chart.api[1].ChartType
        else:
            self.chart_type = chart_type
            self.chart.api[1].ChartType = chart_type

        # self.chart.api[0].Placement = xw.constants.Placement.xlMove
        self.chart.api[0].Placement = Placement.xlFreeFloating
        self.chart.api[0].Border.LineStyle = border_width
        self.chart.api[1].PlotVisibleOnly = visible_only

        self.xaxis.MajorTickMark = TickMark.xlTickMarkInside
        self.yaxis.MajorTickMark = TickMark.xlTickMarkInside

        self.chart.api[1].HasLegend = has_legend
        self.chart.api[1].Legend.IncludeInLayout = include_in_layout

        self.series_collection = []

    @property
    def xaxis(self):  # noqa: ANN201
        chart = self.chart.api[1]
        return chart.Axes(AxisType.xlCategory)

    @property
    def yaxis(self):  # noqa: ANN201
        chart = self.chart.api[1]
        return chart.Axes(AxisType.xlValue)

    def add_series(
        self,
        x: Any,
        y: Any | None = None,
        label: str | tuple[int, int] | Range = "",
        sheet: Sheet | None = None,
        chart_type: int | None = None,
    ) -> Series:
        sheet = sheet or self.sheet

        if chart_type is None:
            chart_type = self.chart_type

        series = Series(self.chart, x, y, label, chart_type, sheet)
        self.series_collection.append(series)

        return series

    @property
    def title(self) -> str | None:
        api = self.chart.api[1]

        if api.HasTitle:
            return api.ChartTitle.Text

        return None

    @title.setter
    def title(self, value: str | tuple[int, int] | Range | None) -> None:
        self.set_title(value)

    def set_title(
        self,
        title: str | tuple[int, int] | Range | None = None,
        *,
        name: str | None = None,
        size: int | None = None,
        sheet: Sheet | None = None,
        **kwargs,
    ) -> None:
        api = self.chart.api[1]

        if title is None:
            api.HasTitle = False
            return

        api.HasTitle = True
        chart_title = api.ChartTitle
        chart_title.Text = reference(title, sheet or self.chart.parent)

        size = size or rcParams["chart.title.font.size"]
        set_font_api(chart_title, name, size=size, **kwargs)

    @property
    def xlabel(self) -> str | None:
        return get_axis_label(self.xaxis)

    @xlabel.setter
    def xlabel(self, value: str | tuple[int, int] | Range | None) -> None:
        self.set_xlabel(value)

    @property
    def ylabel(self) -> str | None:
        return get_axis_label(self.yaxis)

    @ylabel.setter
    def ylabel(self, value: str | tuple[int, int] | Range | None) -> None:
        self.set_ylabel(value)

    def set_xlabel(
        self,
        label: str | tuple[int, int] | Range | None = None,
        sheet: Sheet | None = None,
        **kwargs,
    ) -> None:
        sheet = sheet or self.chart.parent
        set_axis_label(self.xaxis, label, sheet=sheet, **kwargs)

    def set_ylabel(
        self,
        label: str | tuple[int, int] | Range | None = None,
        sheet: Sheet | None = None,
        **kwargs,
    ) -> None:
        sheet = sheet or self.chart.parent
        set_axis_label(self.yaxis, label, sheet=sheet, **kwargs)

    @property
    def xticks(self) -> tuple[float, float, float, float]:
        return get_ticks(self.xaxis)

    @xticks.setter
    def xticks(self, value: tuple[float, ...]) -> None:
        set_ticks(self.xaxis, *value)

    @property
    def yticks(self) -> tuple[float, float, float, float]:
        return get_ticks(self.yaxis)

    @yticks.setter
    def yticks(self, value: tuple[float, ...]) -> None:
        set_ticks(self.yaxis, *value)

    def set_xticks(self, *args, **kwargs) -> None:
        set_ticks(self.xaxis, *args, **kwargs)

    def set_yticks(self, *args, **kwargs) -> None:
        set_ticks(self.yaxis, *args, **kwargs)

    def set_xtick_labels(self, *args, **kwargs) -> None:
        set_tick_labels(self.xaxis, *args, **kwargs)

    def set_ytick_labels(self, *args, **kwargs) -> None:
        set_tick_labels(self.yaxis, *args, **kwargs)

    @property
    def xscale(self) -> str:
        return get_axis_scale(self.xaxis)

    @xscale.setter
    def xscale(self, scale: str) -> None:
        set_axis_scale(self.xaxis, scale)

    @property
    def yscale(self) -> str:
        return get_axis_scale(self.yaxis)

    @yscale.setter
    def yscale(self, scale: str) -> None:
        set_axis_scale(self.yaxis, scale)

    def set(
        self,
        xlabel: str | tuple[int, int] | Range | None = "",
        ylabel: str | tuple[int, int] | Range | None = "",
        xticks: tuple[float, ...] | None = None,
        yticks: tuple[float, ...] | None = None,
        xscale: str | None = None,
        yscale: str | None = None,
    ) -> None:
        if xlabel != "":
            self.xlabel = xlabel
        if ylabel != "":
            self.ylabel = ylabel
        if xticks:
            self.xticks = xticks
        if yticks:
            self.yticks = yticks
        if xscale:
            self.xscale = xscale
        if yscale:
            self.yscale = yscale

    def delete_legend(self) -> None:
        api = self.chart.api[1]
        if api.HasLegend:
            api.Legend.Delete()

    def set_legend(
        self,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        *,
        name: str | None = None,
        size: int | None = None,
        border: str | int = "#aaaaaa",
        fill: str | int = "#eeeeee",
        alpha: float = 0.8,
        loc: tuple[float, float] | None = (1, 1),
        margin: float = 3,
        entry_height_scale: float = 1,
    ) -> None:
        self.delete_legend()
        api = self.chart.api[1]
        api.HasLegend = True

        legend = api.Legend
        legend.IncludeInLayout = False

        legend_entries = list(legend.LegendEntries())
        it = zip(legend_entries, self.series_collection, strict=True)
        for entry, series in it:
            if not series.label:
                entry.Delete()

        size = size or rcParams["chart.legend.font.size"]

        if api.HasLegend is False:
            return

        set_font_api(legend, name, size=size)

        if height is None:
            heights = [0]
            for entry in legend.LegendEntries():
                with suppress(Exception):
                    heights.append(entry.Height * entry_height_scale)
            height = sum(heights)

        if width is None:
            widths = [0]
            for entry in legend.LegendEntries():
                with suppress(Exception):
                    widths.append(entry.Width)
            width = max(widths)

        set_dimensions(legend, left, top, width, height)
        set_area_format(legend, border, fill, alpha)

        if loc:
            x, y = loc
            x = (x + 1) / 2
            y = (1 - y) / 2

            plot_area = api.PlotArea
            inside_left = plot_area.InsideLeft + margin
            inside_top = plot_area.InsideTop + margin
            inside_width = plot_area.InsideWidth - 2 * margin
            inside_height = plot_area.InsideHeight - 2 * margin

            left = inside_left + x * inside_width - x * legend.Width
            top = inside_top + y * inside_height - y * legend.Height

            set_dimensions(legend, left, top)

    def tight_layout(self, title_height_scale: float = 0.7) -> None:
        api = self.chart.api[1]
        xaxis = self.xaxis
        yaxis = self.yaxis

        if not (api.HasTitle and xaxis.HasTitle and yaxis.HasTitle):
            return

        title = api.ChartTitle
        pa = api.PlotArea
        ga = self.chart.api[0]

        title.Top = 0
        yaxis.AxisTitle.Left = 0
        xaxis.AxisTitle.Top = ga.Height - xaxis.AxisTitle.Height

        pa.Top = title_height_scale * title.Height
        pa.Left = yaxis.AxisTitle.Width
        pa.Width = ga.Width - pa.Left
        pa.Height = ga.Height - pa.Top - xaxis.AxisTitle.Height

        title.Left = pa.InsideLeft + pa.InsideWidth / 2 - title.Width / 2

        xaxis.AxisTitle.Left = (
            pa.InsideLeft + pa.InsideWidth / 2 - xaxis.AxisTitle.Width / 2
        )
        yaxis.AxisTitle.Top = (
            pa.InsideTop + pa.InsideHeight / 2 - yaxis.AxisTitle.Height / 2
        )

    def set_plot_area_style(self) -> None:
        # msoElementPrimaryCategoryGridLinesMajor = 334
        api = self.chart.api[1]
        api.SetElement(334)
        # msoElementPrimaryValueGridLinesMajor == 330
        api.SetElement(330)

        plot_area = api.PlotArea
        line = plot_area.Format.Line
        line.Visible = True
        line.ForeColor.RGB = 0
        line.Weight = 1.25
        line.Transparency = 0

        line = self.xaxis.MajorGridlines.Format.Line
        line.Visible = True
        line.ForeColor.RGB = 0
        line.Weight = 1
        line.Transparency = 0.7

        line = self.yaxis.MajorGridlines.Format.Line
        line.Visible = True
        line.ForeColor.RGB = 0
        line.Weight = 1
        line.Transparency = 0.7

    # @property
    # def plot_area(self) -> COMRetryObjectWrapper:
    #     return self.chart.api[1].PlotArea

    # @property
    # def graph_area(self) -> COMRetryObjectWrapper:
    #     return self.chart.api[0]
