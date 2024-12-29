import xlwings as xw

from xlviews.config import rcParams
from xlviews.style import (
    set_area,
    set_dimensions,
    set_font,
    set_label,
    set_scale,
    set_ticklabels,
    set_ticks,
)
from xlviews.utils import constant, multirange, reference

first_position = {"left": 50, "top": 80}


def set_first_position(sf, pos="right"):
    if pos == "right":
        cell = sf.get_adjacent_cell(offset=0)
        first_position["left"] = cell.left
        first_position["top"] = cell.top
    elif pos == "inside":
        cell = sf.cell.offset(sf.columns_level, sf.index_level)
        first_position["left"] = cell.left + 30
        first_position["top"] = cell.top + 30
    elif pos == "bottom":
        cell = sf.cell.offset(sf.columns_level + len(sf) + 1)
        first_position["left"] = cell.left
        first_position["top"] = cell.top


def chart_position(sheet, left, top):
    """
    チャートの位置を自動で設定するための関数．
    left == 0 かつ top is None の場合に，改行する．
    left == None かつ top == Noneのときは右隣に配置する．
    """
    if left is not None and top is not None:
        return left, top
    if not sheet.charts:
        return first_position["left"], first_position["top"]
    elif left == 0 and top is None:  # チャートの改行
        top = 0
        left = 1e100
        for chart in sheet.charts:
            top = max(top, chart.top)
            left = min(left, chart.left)
        for chart in sheet.charts:
            if chart.top == top:
                top = max(top, chart.top + chart.height)
    else:  # チャートの右つなぎ
        chart = sheet.charts[-1]
        left = chart.left + chart.width
        top = chart.top
    return left, top


class Axes(object):
    def __init__(
        self,
        sheet=None,
        left=None,
        top=None,
        width=None,
        height=None,
        row=None,
        column=None,
        visible_only=True,
        border_width=0,
        has_legend=True,
        include_in_layout=False,
    ):
        self.sheet = sheet if sheet else xw.sheets.active

        if row:
            top = self.sheet.range(row, 1).top
        if column:
            left = self.sheet.range(1, column).left

        left, top = chart_position(self.sheet, left, top)
        width = width or rcParams["chart.width"]
        height = height or rcParams["chart.height"]

        self.chart = self.sheet.charts.add(
            left=left, top=top, width=width, height=height
        )
        self.chart_type = None

        # self.chart.api[0].Placement = xw.constants.Placement.xlMove
        self.chart.api[0].Placement = xw.constants.Placement.xlFreeFloating
        self.chart.api[0].Border.LineStyle = border_width
        self.chart.api[1].PlotVisibleOnly = visible_only

        self.xaxis.MajorTickMark = xw.constants.TickMark.xlTickMarkInside
        self.yaxis.MajorTickMark = xw.constants.TickMark.xlTickMarkInside

        self.chart.api[1].HasLegend = has_legend
        self.legend.IncludeInLayout = include_in_layout

        self.series_collection = []
        self.labels = []

    def set_chart_type(self, chart_type):
        """
        チャートの種類を設定する．

        Parameters
        ----------
        chart_type : str
            XYScatterなどの文字列
            xlwings.constants.ChartTypeのメンバを指定する．
            先頭の'xl'は省略する．
        """
        chart = self.chart.api[1]
        self.chart_type = chart_type
        if isinstance(chart_type, str):
            chart_type = constant("ChartType", chart_type)
        chart.ChartType = chart_type

    def add_series(
        self,
        index=None,
        columns=None,
        name=None,
        sheet=None,
        series=None,
        axis=1,
        chart_type=None,
    ):
        """
        シリーズを追加する．

        Parameters
        ----------
        index : list or xlwings.Range
            データインデックス
            See also: xlviews.utils.multirange
        columns : int or list or xlwings.Range, optional
            intの場合，yの値のみ，listの場合(x, y)の値
        name : tuple or str
            tuple の場合，(row, col)
        sheet : str
            データソースのあるシート名
        series : Excelチャートのシリーズ
            指定したばあい，すでに存在するシリーズを変更する．
        axis : int
            データの方向
        chart_type : int or str, optional
            チャートタイプ

        Returns
        -------
        series : Series
            シリーズオブジェクト
        """
        if sheet is None:
            sheet = self.sheet
        if series is None:
            chart = self.chart.api[1]
            series = chart.SeriesCollection().NewSeries()
            self.series_collection.append(series)
            self.labels.append(name)
        if name:
            name = reference(sheet, name)
            series.Name = name

        def _multirange(index_, column):
            if axis == 1:
                values = multirange(sheet, index_, column)
            else:
                values = multirange(sheet, column, index_)
            return values

        if not isinstance(index, list) and columns is not None:
            series.XValues = index.api
            series.Values = columns.api
        elif isinstance(columns, int):
            series.Values = _multirange(index, columns)
        elif isinstance(columns, list):
            series.XValues = _multirange(index, columns[0])
            series.Values = _multirange(index, columns[1])
        else:
            raise ValueError("columnsが指定されていない．")

        if chart_type and self.chart_type != chart_type:
            if isinstance(chart_type, str):
                chart_type = constant("ChartType", chart_type)
            series.ChartType = chart_type

        return series

    def set_title(self, title=None, name=None, size=None, sheet=None, **kwargs):
        """
        チャートのタイトルを設定する．

        Parameters
        ----------
        title : str or list or range
            文字列で直接設定するか，[row, column]の参照
        name : str
            フォント名
        size : int
            文字サイズ
        sheet : シートオブジェクト
            セル参照するときのシート
        """
        chart = self.chart.api[1]
        if sheet is None:
            sheet = self.chart.parent
        if title is None:
            chart.HasTitle = False
            return
        chart.HasTitle = True
        chart_title = chart.ChartTitle
        chart_title.Text = reference(sheet, title)
        if size is None:
            size = rcParams["chart.title.font.size"]
        set_font(chart_title, name=name, size=size, **kwargs)

    def set_legend(
        self,
        legend=True,
        name=None,
        size=None,
        left=None,
        top=None,
        width=None,
        height=None,
        fill="yellow",
        border="gray",
        alpha=0.8,
        position=(1, 1),
        margin=3,
        entry_height_scale=1,
    ):
        if self.chart.api[1].HasLegend:
            self.legend.Delete()
        if not legend:
            return
        else:
            # 表示されないLegendEntryのHeightやWidthを取得できないため
            self.chart.api[1].HasLegend = True
            self.legend.IncludeInLayout = False

            legend_entries = list(self.legend.LegendEntries())
            labels = [label for label in self.labels if label != "__trendline__"]
            labels += [None for label in self.labels if label == "__trendline__"]
            for entry, label in zip(legend_entries, labels):
                if label is None:
                    entry.Delete()

        if size is None:
            size = rcParams["chart.legend.font.size"]
        # ここでチェックしないとだめ
        if self.chart.api[1].HasLegend is False:
            return
        set_font(self.legend, name=name, size=size)

        # TODO: 凡例が一列以外の場合
        if height is None:
            heights = [0]
            for entry in self.legend.LegendEntries():
                try:
                    heights.append(entry.Height * entry_height_scale)
                except Exception:
                    pass
            height = sum(heights)
        if width is None:
            widths = [0]
            for entry in self.legend.LegendEntries():
                try:
                    widths.append(entry.Width)
                except Exception:
                    pass
            width = max(widths)

        set_dimensions(self.legend, left, top, width, height)
        set_area(self.legend, fill=fill, border=border, alpha=alpha)

        if position:
            legend = self.legend
            plot_area = self.plot_area
            x, y = position
            x = (x + 1) / 2
            y = (1 - y) / 2

            # マージン分だけInsideAreaを縮小する．
            inside_left = plot_area.InsideLeft + margin
            inside_top = plot_area.InsideTop + margin
            inside_width = plot_area.InsideWidth - 2 * margin
            inside_height = plot_area.InsideHeight - 2 * margin

            left = inside_left + x * inside_width - x * legend.Width
            top = inside_top + y * inside_height - y * legend.Height
            set_dimensions(self.legend, left, top)

    def get_axis(self, axis):
        chart = self.chart.api[1]
        if axis == "x":
            return chart.Axes(xw.constants.AxisType.xlCategory)
        elif axis == "y":
            return chart.Axes(xw.constants.AxisType.xlValue)

    @property
    def xaxis(self):
        return self.get_axis("x")

    @property
    def yaxis(self):
        return self.get_axis("y")

    @property
    def title(self):
        return self.chart.api[1].ChartTitle

    @property
    def plot_area(self):
        return self.chart.api[1].PlotArea

    @property
    def graph_area(self):
        return self.chart.api[0]

    @property
    def legend(self):
        return self.chart.api[1].Legend

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

    def tight_layout(self, title_height_scale=0.7):
        # TODO : タイトル，軸ラベルがない場合でもtight_layout可能にする．
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
        # Major罫線に線を書く．
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
