import pytest
from xlwings import Sheet

from xlviews.chart.axes import Axes
from xlviews.config import rcParams
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")

LEFT = rcParams["chart.left"]
TOP = rcParams["chart.top"]


def test_chart_position_none(sheet: Sheet):
    axes = Axes(sheet=sheet)
    assert axes.chart.left == LEFT
    assert axes.chart.top == TOP


def test_chart_position_from_cell(sheet: Sheet):
    axes = Axes(5, 10, sheet=sheet)
    assert axes.chart.left == 9 * sheet.cells(1, 1).width
    assert axes.chart.top == 4 * sheet.cells(1, 1).height


def test_chart_position_right(sheet: Sheet):
    a = Axes(sheet=sheet)
    b = Axes(sheet=sheet)
    assert b.chart.left == LEFT + a.chart.width
    assert a.chart.top == TOP


def test_chart_position_bottom(sheet: Sheet):
    a = Axes(sheet=sheet)
    b = Axes(sheet=sheet, top=-1)
    assert b.chart.left == LEFT
    assert b.chart.top == TOP + a.chart.height


def test_chart_position_left(sheet: Sheet):
    a = Axes(sheet=sheet)
    Axes(sheet=sheet)
    c = Axes(sheet=sheet, left=0)
    assert c.chart.left == LEFT
    assert c.chart.top == a.chart.top + a.chart.height


def test_chart_position_left_top(sheet: Sheet):
    Axes(sheet=sheet)
    Axes(sheet=sheet)
    c = Axes(sheet=sheet, left=0, top=100)
    assert c.chart.left == LEFT
    assert c.chart.top == 100


def test_chart_position_top(sheet: Sheet):
    a = Axes(sheet=sheet)
    Axes(sheet=sheet, top=-1)
    c = Axes(sheet=sheet, top=0)
    assert c.chart.left == a.chart.left + a.chart.width
    assert c.chart.top == TOP


def test_chart_position_top_left(sheet: Sheet):
    Axes(sheet=sheet)
    Axes(sheet=sheet, top=-1)
    c = Axes(sheet=sheet, top=0, left=200)
    assert c.chart.left == 200
    assert c.chart.top == TOP


def test_copy_right(sheet: Sheet):
    a = Axes(sheet=sheet)
    b = a.copy(left=0)
    assert b.chart.left == a.chart.left + a.chart.width


def test_copy_bottom(sheet: Sheet):
    a = Axes(sheet=sheet)
    b = a.copy(top=0)
    assert b.chart.top == a.chart.top + a.chart.height
