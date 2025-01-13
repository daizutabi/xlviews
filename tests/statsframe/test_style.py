import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame
from xlviews.utils import rgb


@pytest.fixture
def sf(df: DataFrame, sheet: Sheet):
    return SheetFrame(3, 3, data=df, table=True, sheet=sheet)


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
    sf = sf.statsframe(func, by="x", table=True)
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
