import pytest
from xlwings import Sheet
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.chart.series import Series
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def ax(sheet_module: Sheet):
    ct = ChartType.xlXYScatterLines
    return Axes(300, 10, chart_type=ct, sheet=sheet_module)


@pytest.fixture(scope="module")
def series(ax: Axes):
    return ax.add_series(
        [1, 2, 3],
        [4, 5, 6],
        label="xy",
        chart_type=ChartType.xlXYScatterLines,
    )


@pytest.fixture
def series_y(ax: Axes):
    series = ax.add_series([7, 8, 9], label="y")
    yield series
    series.delete()


def test_series_x(series: Series):
    assert series.x == ("1", "2", "3")


def test_series_y(series: Series):
    assert series.y == (4, 5, 6)


def test_series_y_y(series_y: Series):
    assert series_y.y == (7, 8, 9)


def test_chart_type(series: Series):
    assert series.chart_type == ChartType.xlXYScatterLines


def test_name(series: Series):
    assert series.label == "xy"
