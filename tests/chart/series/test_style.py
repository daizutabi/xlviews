import pytest
from xlwings import Sheet
from xlwings.constants import ChartType, LineStyle, MarkerStyle

from xlviews.chart.axes import Axes
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def x(sheet_module: Sheet):
    x = sheet_module["B2:B11"]
    x.options(transpose=True).value = list(range(10))
    return x


@pytest.fixture(scope="module")
def y(sheet_module: Sheet):
    y = sheet_module["C2:C11"]
    y.options(transpose=True).value = list(range(10, 20))
    return y


@pytest.fixture(scope="module")
def ax(sheet_module: Sheet):
    ct = ChartType.xlXYScatterLines
    return Axes(300, 10, chart_type=ct, sheet=sheet_module)


@pytest.mark.parametrize(
    ("style", "value", "size"),
    [
        ("", MarkerStyle.xlMarkerStyleNone, 5),
        ("o", MarkerStyle.xlMarkerStyleCircle, 10),
        ("^", MarkerStyle.xlMarkerStyleTriangle, 9),
        ("s", MarkerStyle.xlMarkerStyleSquare, 8),
        ("d", MarkerStyle.xlMarkerStyleDiamond, 7),
        ("+", MarkerStyle.xlMarkerStylePlus, 6),
        ("x", MarkerStyle.xlMarkerStyleX, 5),
        (".", MarkerStyle.xlMarkerStyleDot, 4),
        ("-", MarkerStyle.xlMarkerStyleDash, 3),
        ("*", MarkerStyle.xlMarkerStyleStar, 2),
    ],
)
def test_series_style_marker(ax: Axes, x, y, style, value, size):
    series = ax.add_series(x, y, label="a")
    series.marker(style, size=size)
    assert series.api.MarkerStyle == value
    assert series.api.MarkerSize == size
    series.delete()


@pytest.mark.parametrize(
    ("style", "value", "weight"),
    [
        ("", LineStyle.xlLineStyleNone, 2),
        ("-", LineStyle.xlContinuous, 1),
        ("--", LineStyle.xlDash, 2),
        ("-.", LineStyle.xlDashDot, 1),
        (".", LineStyle.xlDot, 2),
    ],
)
def test_series_style_line(ax: Axes, x, y, style, value, weight):
    series = ax.add_series(x, y, label="a")
    series.line(style, weight=weight)
    assert series.api.Border.LineStyle == value
    series.delete()


def test_series_style_line_none(x, y, sheet: Sheet):
    ct = ChartType.xlXYScatterLines
    ax = Axes(300, 10, chart_type=ct, sheet=sheet)
    series = ax.add_series(x, y, label="a")
    style = series.api.Border.LineStyle
    series.line()
    assert series.api.Border.LineStyle == style


@pytest.mark.parametrize(
    ("color", "value", "alpha"),
    [
        ("red", 255, 0.5),
        ("green", 32768, 0.2),
        ("blue", 16711680, 0.4),
        ("lime", 65280, 0.7),
    ],
)
def test_series_style_color(ax: Axes, x, y, color, value, alpha):
    series = ax.add_series(x, y, label="a")
    series.line("-", color=color, alpha=alpha)
    assert value == series.api.Format.Line.ForeColor.RGB
    series.delete()
