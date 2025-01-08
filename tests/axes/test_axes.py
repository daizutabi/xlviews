import pytest
from xlwings import Sheet

# from xlviews import constant
from xlviews.axes import Axes

# from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def axes(sheet_module: Sheet):
    return Axes(sheet=sheet_module)


# @pytest.mark.parametrize("axis", ["xaxis", "yaxis"])
# def test_xaxis(axes: Axes, axis: str):
#     mark = constant("TickMark", "TickMarkInside")
#     assert getattr(axes, axis).MajorTickMark == mark


# @pytest.mark.parametrize(
#     ("chart_type_str", "chart_type_int"),
#     [("XYScatter", 74), ("Line", 65), ("XYScatterLines", 74)],
# )
# def test_chart_type(axes: Axes, chart_type_str: str, chart_type_int: int):
#     axes.set_chart_type(chart_type_str)
#     assert axes.chart_type == chart_type_str
#     assert axes.chart.api[1].ChartType == chart_type_int


# def test_add_series(axes: Axes):
#     pass


if __name__ == "__main__":
    import xlwings as xw

    from xlviews.axes import Axes
    from xlviews.common import quit_apps

    quit_apps()
    book = xw.Book()
    sheet = book.sheets.add()
    axes = Axes(sheet=sheet)
    # print(1)
