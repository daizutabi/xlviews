import pytest
from xlwings import Sheet
from xlwings.constants import ChartType, TickMark

from xlviews.axes import Axes


@pytest.fixture
def ax(sheet_module: Sheet):
    ct = ChartType.xlXYScatterLines
    return Axes(300, 10, chart_type=ct, sheet=sheet_module)


def test_chart_position(ax: Axes):
    assert ax.chart.left == 300
    assert ax.chart.top == 9.75


def test_chart_dimensions(ax: Axes):
    assert ax.chart.width == 200
    assert ax.chart.height == 199.5


def test_chart_type(ax: Axes):
    assert ax.chart.api[1].ChartType == ChartType.xlXYScatterLines
    assert ax.chart.chart_type == "xy_scatter_lines"


def test_chart_type_none(sheet_module: Sheet):
    ax = Axes(sheet=sheet_module)
    assert ax.chart.api[1].ChartType == ChartType.xlColumnClustered
    assert ax.chart.chart_type == "column_clustered"


def test_chart_row_column(sheet_module: Sheet):
    ax = Axes(row=2, column=3, sheet=sheet_module)
    assert ax.chart.left == 104
    assert ax.chart.top == 18


@pytest.mark.parametrize("axis", ["xaxis", "yaxis"])
def test_axis(ax: Axes, axis: str):
    mark = TickMark.xlTickMarkInside
    assert getattr(ax, axis).MajorTickMark == mark
