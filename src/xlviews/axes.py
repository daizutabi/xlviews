from __future__ import annotations

from contextlib import suppress
from typing import TYPE_CHECKING

import xlwings as xw
from xlwings import Range
from xlwings.constants import AxisType, Placement, TickMark

from xlviews.config import rcParams
from xlviews.range import reference
from xlviews.style import (
    set_area_format,
    set_dimensions,
    set_font_api,
    set_label,
    set_scale,
    set_ticklabels,
    set_ticks,
)

if TYPE_CHECKING:
    from xlwings import Chart, Sheet
    from xlwings._xlwindows import COMRetryObjectWrapper

    from xlviews.frame import SheetFrame

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


def chart_position(sheet: Sheet, left: int | None, top: int | None) -> tuple[int, int]:
    """Return the position of the chart.

    If left is 0 and top is None, it will create a new row.
    If left is None and top is None, it will be placed to the right.
    """
    if left is not None and top is not None:
        return left, top

    if not sheet.charts:
        return FIRST_POSITION["left"], FIRST_POSITION["top"]

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
    series_collection: list[COMRetryObjectWrapper]
    labels: list[str]

    def __init__(
        self,
        left: int | None = None,
        top: int | None = None,
        width: int = 0,
        height: int = 0,
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

        self.chart = self.sheet.charts.add(left, top, width, height)

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
        self.labels = []

    @property
    def xaxis(self) -> COMRetryObjectWrapper:
        chart = self.chart.api[1]
        return chart.Axes(AxisType.xlCategory)

    @property
    def yaxis(self) -> COMRetryObjectWrapper:
        chart = self.chart.api[1]
        return chart.Axes(AxisType.xlValue)

    def add_series(
        self,
        x: Range,
        y: Range | None = None,
        label: str | tuple[int, int] | Range = "",
        sheet: Sheet | None = None,
        chart_type: int | None = None,
    ) -> COMRetryObjectWrapper:
        sheet = sheet or self.sheet

        api = self.chart.api[1]
        series = api.SeriesCollection().NewSeries()
        self.series_collection.append(series)

        if not isinstance(label, str):
            label = reference(label, sheet)

        series.Name = label
        self.labels.append(label)

        if chart_type is None:
            chart_type = self.chart_type

        series.ChartType = chart_type

        if y:
            series.XValues = x.api
            series.Values = y.api

        else:
            series.Values = x.api

        return series

    @property
    def title(self) -> str | None:
        api = self.chart.api[1]

        if api.HasTitle:
            return api.ChartTitle.Text

        return None

    @title.setter
    def title(self, value: str | tuple[int, int] | None) -> None:
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

        sheet = sheet or self.chart.parent

        api.HasTitle = True
        chart_title = api.ChartTitle
        chart_title.Text = reference(title, sheet)

        size = size or rcParams["chart.title.font.size"]
        set_font_api(chart_title, name, size=size, **kwargs)

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
        border: str | int = "gray",
        fill: str | int = "yellow",
        alpha: float = 0.8,
        position: tuple[float, float] | None = (1, 1),
        margin: float = 3,
        entry_height_scale: float = 1,
    ) -> None:
        self.delete_legend()
        api = self.chart.api[1]
        api.HasLegend = True

        legend = api.Legend
        legend.IncludeInLayout = False

        legend_entries = list(legend.LegendEntries())
        for entry, label in zip(legend_entries, self.labels, strict=True):
            if not label:
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

        if position:
            x, y = position
            x = (x + 1) / 2
            y = (1 - y) / 2

            plot_area = self.plot_area
            inside_left = plot_area.InsideLeft + margin
            inside_top = plot_area.InsideTop + margin
            inside_width = plot_area.InsideWidth - 2 * margin
            inside_height = plot_area.InsideHeight - 2 * margin

            left = inside_left + x * inside_width - x * legend.Width
            top = inside_top + y * inside_height - y * legend.Height

            set_dimensions(legend, left, top)

    def set_xscale(self, scale=None, **kwargs):
        set_scale(self.xaxis, scale, **kwargs)

    def set_yscale(self, scale=None, **kwargs):
        set_scale(self.yaxis, scale, **kwargs)

    def set_xlabel(self, label=None, **kwargs):
        set_label(self.xaxis, label, **kwargs)

    def set_ylabel(self, label=None, **kwargs):
        set_label(self.yaxis, label, **kwargs)

    def set_xticks(self, *args, **kwargs):
        set_ticks(self.xaxis, *args, **kwargs)

    def set_yticks(self, *args, **kwargs):
        set_ticks(self.yaxis, *args, **kwargs)

    def set_xticklabels(self, *args, **kwargs):
        set_ticklabels(self.xaxis, *args, **kwargs)

    def set_yticklabels(self, *args, **kwargs):
        set_ticklabels(self.yaxis, *args, **kwargs)

    @property
    def plot_area(self):
        return self.chart.api[1].PlotArea

    @property
    def graph_area(self):
        return self.chart.api[0]

    def tight_layout(self, title_height_scale=0.7):
        # TODO : タイトル、軸ラベルがない場合でもtight_layout可能にする。
        if not (
            self.chart.api[1].HasTitle and self.xaxis.HasTitle and self.yaxis.HasTitle
        ):
            return

        self.title.Top = 0
        self.yaxis.AxisTitle.Left = 0
        self.xaxis.AxisTitle.Top = self.graph_area.Height - self.xaxis.AxisTitle.Height
        self.plot_area.Top = title_height_scale * self.title.Height
        self.plot_area.Left = self.yaxis.AxisTitle.Width
        self.plot_area.Width = self.graph_area.Width - self.plot_area.Left - 0
        self.plot_area.Height = (
            self.graph_area.Height
            - self.plot_area.Top
            - self.xaxis.AxisTitle.Height
            - 0
        )

        self.title.Left = (
            self.plot_area.InsideLeft
            + self.plot_area.InsideWidth / 2
            - self.title.Width / 2
        )

        self.xaxis.AxisTitle.Left = (
            self.plot_area.InsideLeft
            + self.plot_area.InsideWidth / 2
            - self.xaxis.AxisTitle.Width / 2
        )
        self.yaxis.AxisTitle.Top = (
            self.plot_area.InsideTop
            + self.plot_area.InsideHeight / 2
            - self.yaxis.AxisTitle.Height / 2
        )

    def set_plot_area_style(self):
        # Major罫線に線を書く。
        # msoElementPrimaryCategoryGridLinesMajor = 334
        self.chart.api[1].SetElement(334)
        # msoElementPrimaryValueGridLinesMajor == 330
        self.chart.api[1].SetElement(330)

        line = self.plot_area.Format.Line
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
