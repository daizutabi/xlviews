import numpy as np
import pytest

from xlviews.colors import rgb
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.stats_frame import StatsFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def funcs():
    return ["count", "mean", "median", "min", "max", "soa", "sum", "std"]


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame, funcs: list[str]):
    return StatsFrame(sf_parent, funcs, by=":y")


def test_func(sf: StatsFrame, funcs: list[str]):
    rng = sf.sheet.range((sf.row + 1, sf.column), (sf.row + len(sf), sf.column))
    assert rng.value == np.tile(funcs, 4).tolist()


@pytest.mark.parametrize(
    ("func", "color"),
    [
        ("count", "gray"),
        ("mean", "#33aa33"),
        ("median", "#111111"),
        ("min", "#7777FF"),
        ("max", "#FF7777"),
        ("soa", "#5555FF"),
        ("sum", "purple"),
        ("std", "#aaaaaa"),
    ],
)
@pytest.mark.parametrize("c", ["func", "c"])
@pytest.mark.parametrize("o", [0, 8])
def test_color(sf: StatsFrame, funcs: list[str], func: str, color, c, o):
    row = funcs.index(func) + sf.row + o + 1
    column = sf.get_loc(c)
    rng = sf.sheet.range(row, column)
    assert rgb(rng.font.color) == rgb(color)


@pytest.mark.parametrize("func", ["soa", "sum"])
@pytest.mark.parametrize("c", ["func", "a"])
@pytest.mark.parametrize("o", [16, 24])
def test_italic(sf: StatsFrame, funcs: list[str], func: str, c, o):
    row = funcs.index(func) + sf.row + o + 1
    column = sf.get_loc(c)
    rng = sf.sheet.range(row, column)
    assert rng.font.italic


@pytest.mark.parametrize("c", ["a", "b"])
def test_soa(sf: StatsFrame, funcs: list[str], c):
    row = funcs.index("soa") + sf.row + 1
    column = sf.get_loc(c)
    rng = sf.sheet.range(row, column)
    assert rng.number_format == "0.0%"


@pytest.mark.parametrize("func", ["median", "min", "mean", "max", "std", "sum"])
@pytest.mark.parametrize("o", [0, 16])
def test_number_format(sf: StatsFrame, funcs: list[str], func, o):
    row = funcs.index(func) + sf.row + o + 1
    column = sf.get_loc("c")
    rng = sf.sheet.range(row, column)
    assert rng.number_format == "0.00"
