import pytest
from xlwings import Sheet
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture
def ax(sheet: Sheet):
    ct = ChartType.xlXYScatterLines
    return Axes(2, 2, ct, sheet)


def test_add_series_xy(ax: Axes):
    x = ax.sheet.range("A1:A10")
    y = ax.sheet.range("B1:B10")

    x.options(transpose=True).value = list(range(10))
    y.options(transpose=True).value = list(range(10, 20))
    s = ax.add_series(x, y)

    assert s.api.XValues == (0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    assert s.x == (0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    assert s.api.Values == (10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
    assert s.y == (10, 11, 12, 13, 14, 15, 16, 17, 18, 19)

    x.options(transpose=True).value = list(range(20, 30))
    y.options(transpose=True).value = list(range(30, 40))

    assert s.api.XValues == (20, 21, 22, 23, 24, 25, 26, 27, 28, 29)
    assert s.x == (20, 21, 22, 23, 24, 25, 26, 27, 28, 29)
    assert s.api.Values == (30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
    assert s.y == (30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
    assert s.api.ChartType == ChartType.xlXYScatterLines
    assert s.chart_type == ChartType.xlXYScatterLines


def test_add_series_x(ax: Axes):
    x = ax.sheet.range("C1:C5")
    x.options(transpose=True).value = list(range(100, 105))
    s = ax.add_series(x)
    assert s.api.XValues == ("1", "2", "3", "4", "5")
    assert s.x == ("1", "2", "3", "4", "5")
    assert s.api.Values == (100, 101, 102, 103, 104)
    assert s.y == (100, 101, 102, 103, 104)


def test_add_series_chart_type(ax: Axes):
    x = ax.sheet.range("D1:D5")
    s = ax.add_series(x, chart_type=ChartType.xlXYScatter)
    assert s.api.ChartType == ChartType.xlXYScatter
    assert s.chart_type == ChartType.xlXYScatter


def test_add_series_name_range(ax: Axes):
    x = ax.sheet.range("A2:A5")
    rng = Range(1, 1, ax.sheet)
    label = rng.get_address(include_sheetname=True, formula=True)
    s = ax.add_series(x, label=label)

    rng.value = "Series Name"
    assert s.api.Name == "Series Name"
    assert s.label == "Series Name"


def test_add_series_xy_range_collection(ax: Axes):
    from xlviews.core.range_collection import RangeCollection

    ax.sheet.range("A1:A10").options(transpose=True).value = list(range(10))
    ax.sheet.range("B1:B10").options(transpose=True).value = list(range(10, 20))

    x = RangeCollection([(1, 3), (8, 10)], 1, ax.sheet)
    y = RangeCollection([(1, 3), (8, 10)], 2, ax.sheet)
    s = ax.add_series(x, y)

    assert s.api.XValues == (0, 1, 2, 7, 8, 9)
    assert s.x == (0, 1, 2, 7, 8, 9)
    assert s.api.Values == (10, 11, 12, 17, 18, 19)
    assert s.y == (10, 11, 12, 17, 18, 19)
