import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame
from xlviews.stats import StatsFrame
from xlviews.utils import rgb


@pytest.fixture
def sf(df: DataFrame, sheet: Sheet):
    return SheetFrame(sheet, 3, 3, data=df, table=True)


@pytest.mark.parametrize(
    ("func", "color"),
    [
        ("count", "gray"),
        ("max", "#FF7777"),
        ("mean", "#33aa33"),
        ("min", "#7777FF"),
        ("soa", "#5555FF"),
        ("std", "#aaaaaa"),
        ("sum", "purple"),
    ],
)
def test_value_style(sf: SheetFrame, func, color):
    sf = StatsFrame(sf, func, by="x", table=True)
    sf.set_value_style("func")
    for c in ["a", "b", "c"]:
        rng = sf.range(c)
        assert rgb(rng.font.color) == rgb(color)
        if func in ["soa", "sum"]:
            assert rng.font.italic
        else:
            assert not rng.font.italic
        if func == "soa":
            assert rng.number_format == "0.0%"
