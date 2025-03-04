import pytest
from xlwings import Sheet
from xlwings.constants import ChartType, MarkerStyle

from xlviews.chart.axes import Axes
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.chart import Base

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize(
    ("label", "key", "expected"),
    [
        ("a", {}, "a"),
        ("a{b}", {"b": "B"}, "aB"),
        ("{a}{b}", {"a": "A", "b": "B"}, "AB"),
        (lambda x: f"_{x['a']}_", {"a": "A"}, "_A_"),
    ],
)
def test_get_label(label, key, expected):
    from xlviews.figure.plot import get_label

    assert get_label(label, key) == expected


def test_get_label_error():
    from xlviews.figure.plot import get_label

    with pytest.raises(ValueError, match="Invalid label"):
        get_label(1, 1)  # type: ignore


@pytest.fixture
def sf(sheet: Sheet):
    fc = Base(sheet, style=True)
    return fc.sf


@pytest.fixture
def ax(sf: SheetFrame):
    return Axes(2, 8, sheet=sf.sheet)


def test_plot_series(ax: Axes, sf: SheetFrame):
    from xlviews.figure.plot import Plot

    data = sf.agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatter)
        .set(label="abc", marker="o", color="blue", alpha=0.6)
    )
    s = p.series_collection[0]
    assert s.label == "abc"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleCircle


def test_plot_index(ax: Axes, sf: SheetFrame):
    from xlviews.figure.plot import Plot

    data = sf.groupby("b").agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatterLines)
        .set(label="b={b}", marker=["o", "s"], color={"s": "red", "t": "blue"})
    )
    assert len(p.series_collection) == 2
    s = p.series_collection[0]
    assert s.label == "b=s"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleCircle
    s = p.series_collection[1]
    assert s.label == "b=t"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleSquare


def test_plot_multi_index(ax: Axes, sf: SheetFrame):
    from xlviews.figure.plot import Plot

    data = sf.groupby(["b", "c"]).agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatterLinesNoMarkers)
        .set(
            label=lambda x: f"{x['b']},{x['c']}",
            marker="b",
            color=("c", ["red", "green"]),
            size=10,
        )
    )
    assert len(p.series_collection) == 4
    s = p.series_collection[0]
    assert s.label == "s,100"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleCircle
    s = p.series_collection[1]
    assert s.label == "s,200"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleCircle
    s = p.series_collection[2]
    assert s.label == "t,100"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleTriangle
    s = p.series_collection[3]
    assert s.label == "t,200"
    assert s.api.MarkerStyle == MarkerStyle.xlMarkerStyleTriangle
