from pathlib import Path

import pytest
from xlwings import Sheet

from xlviews import constant
from xlviews.axes import Axes


def test_set_first_position():
    from xlviews.axes import FIRST_POSITION, set_first_position

    left_org = FIRST_POSITION["left"]
    top_org = FIRST_POSITION["top"]
    assert left_org == 50
    assert top_org == 80

    # TODO: test_set_first_position

    FIRST_POSITION["left"] = left_org
    FIRST_POSITION["top"] = top_org


@pytest.mark.parametrize(
    ("args", "expected"),
    [((10, 20), (10, 20)), ((None, None), (50, 80)), ((0, None), (50, 80))],
)
def test_chart_position(sheet: Sheet, args, expected):
    from xlviews.axes import chart_position

    assert chart_position(sheet, *args) == expected


def test_chart_position_from_cell(sheet: Sheet):
    axes = Axes(sheet=sheet, row=5, column=10)
    assert axes.chart.left == 9 * sheet.cells(1, 1).width
    assert axes.chart.top == 4 * sheet.cells(1, 1).height


def test_chart_position_from_chart(sheet: Sheet):
    a = Axes(sheet=sheet)
    assert a.chart.left == 50

    b = Axes(sheet=sheet)
    assert b.chart.left == a.chart.left + a.chart.width

    c = Axes(0, None, width=200, sheet=sheet)
    assert c.chart.left == a.chart.left
    assert c.chart.top == a.chart.top + a.chart.height

    d = Axes(sheet=sheet)
    assert d.chart.left == c.chart.left + 200
    assert d.chart.top == c.chart.top


@pytest.fixture
def axes(sheet: Sheet):
    return Axes(sheet=sheet)


@pytest.mark.parametrize("axis", ["xaxis", "yaxis"])
def test_xaxis(axes: Axes, axis: str):
    mark = constant("TickMark", "TickMarkInside")
    assert getattr(axes, axis).MajorTickMark == mark


@pytest.mark.parametrize(
    ("chart_type_str", "chart_type_int"),
    [("XYScatter", 74), ("Line", 65), ("XYScatterLines", 74)],
)
def test_chart_type(axes: Axes, chart_type_str: str, chart_type_int: int):
    axes.set_chart_type(chart_type_str)
    assert axes.chart_type == chart_type_str
    assert axes.chart.api[1].ChartType == chart_type_int


def test_add_series(axes: Axes):
    pass
