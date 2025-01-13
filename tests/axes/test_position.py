import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.axes import Axes
from xlviews.sheetframe import SheetFrame


@pytest.mark.parametrize(
    ("pos", "left", "top"),
    [("right", 269.5, 18), ("inside", 134, 66), ("bottom", 52, 90)],
)
def test_set_first_position(sheet: Sheet, pos: str, left: float, top: float):
    from xlviews.axes import FIRST_POSITION, clear_first_position, set_first_position

    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(2, 2, data=df, style=False, autofit=False, sheet=sheet)

    set_first_position(sf, pos)
    assert FIRST_POSITION["left"] == left
    assert FIRST_POSITION["top"] == top
    clear_first_position()


@pytest.mark.parametrize(
    ("args", "expected"),
    [((10, 20), (10, 20)), ((None, None), (50, 50)), ((0, None), (50, 50))],
)
def test_chart_position(sheet: Sheet, args, expected):
    from xlviews.axes import chart_position

    assert chart_position(sheet, *args) == expected


def test_chart_position_from_cell(sheet: Sheet):
    axes = Axes(sheet=sheet, row=5, column=10)
    assert axes.chart.left == 9 * sheet.cells(1, 1).width
    assert axes.chart.top == 4 * sheet.cells(1, 1).height


def test_chart_position_from_chart(sheet: Sheet):
    from xlviews.axes import clear_first_position

    clear_first_position()

    a = Axes(sheet=sheet)
    assert a.chart.left == 50
    assert abs(a.chart.top - 50) <= 0.5

    b = Axes(sheet=sheet)
    assert b.chart.left == a.chart.left + a.chart.width

    c = Axes(0, None, width=200, sheet=sheet)
    assert c.chart.left == a.chart.left
    assert abs(c.chart.top - (a.chart.top + a.chart.height)) <= 0.5

    d = Axes(sheet=sheet)
    assert abs(d.chart.left - (c.chart.left + 200)) <= 0.5
    assert abs(d.chart.top - c.chart.top) <= 0.5
