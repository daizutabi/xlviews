import pytest
from xlwings import Sheet
from xlwings.constants import ChartType

from xlviews.axes import Axes


@pytest.fixture
def ax(sheet_module: Sheet):
    ct = ChartType.xlXYScatterLines
    ax = Axes(300, 10, chart_type=ct, sheet=sheet_module)


if __name__ == "__main__":
    import xlwings as xw
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.common import quit_apps

    quit_apps()
    book = xw.Book()
    sheet = book.sheets.add()
    x = sheet["B2:B11"]
    y = sheet["C2:C11"]
    x.options(transpose=True).value = list(range(10))
    y.options(transpose=True).value = list(range(10, 20))

    ax = Axes(300, 10, sheet=sheet, chart_type=ChartType.xlXYScatterLines)
    ax.add_series(x, y, label=(1, 1))
    ax.title = (2, 1)
    sheet["A1"].value = "abcdef"
    sheet["A2"].value = "title"
    ax.add_series(x, y)

    ax.set_legend(name="abc", position=(-1, -1))
