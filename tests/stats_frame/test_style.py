import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame
from xlviews.utils import is_excel_installed, rgb

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


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
    sf = sf.stats_frame(func, by="x", table=True)
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
