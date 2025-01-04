import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews import constant
from xlviews.axes import Axes
from xlviews.frame import SheetFrame

# @pytest.mark.parametrize(
#     ("pos", "left", "top"),
#     [("right", 112.0, 18.5), ("inside", 94.5, 66.5), ("bottom", 52.0, 90)],
# )
# def test_set_first_position(sheet: Sheet, pos: str, left: float, top: float):
#     from xlviews.axes import FIRST_POSITION, clear_first_position, set_first_position

#     df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False, autofit=False)

#     set_first_position(sf, pos)
#     assert FIRST_POSITION["left"] == left
#     assert FIRST_POSITION["top"] == top
#     clear_first_position()


# @pytest.mark.parametrize(
#     ("args", "expected"),
#     [((10, 20), (10, 20)), ((None, None), (50, 80)), ((0, None), (50, 80))],
# )
# def test_chart_position(sheet: Sheet, args, expected):
#     from xlviews.axes import chart_position

#     assert chart_position(sheet, *args) == expected


# def test_chart_position_from_cell(sheet: Sheet):
#     axes = Axes(sheet=sheet, row=5, column=10)
#     assert axes.chart.left == 9 * sheet.cells(1, 1).width
#     assert axes.chart.top == 4 * sheet.cells(1, 1).height


# def test_chart_position_from_chart(sheet: Sheet):
#     a = Axes(sheet=sheet)
#     assert a.chart.left == 50

#     b = Axes(sheet=sheet)
#     assert b.chart.left == a.chart.left + a.chart.width

#     c = Axes(0, None, width=200, sheet=sheet)
#     assert c.chart.left == a.chart.left
#     assert c.chart.top == a.chart.top + a.chart.height

#     d = Axes(sheet=sheet)
#     assert d.chart.left == c.chart.left + 200
#     assert d.chart.top == c.chart.top


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
