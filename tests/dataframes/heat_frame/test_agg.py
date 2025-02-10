import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_excel_installed
from xlviews.testing.heat_frame.agg import Agg

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Agg(sheet_module)


@pytest.fixture(
    scope="module",
    params=["mean", "A1", (1, 1)],
    ids=["mean", "A1", "Range"],
)
def sf(fc: Agg, request: pytest.FixtureRequest):
    fc.sf.sheet.range("A1").value = "mean"
    aggfunc = request.param
    if isinstance(aggfunc, tuple):
        aggfunc = fc.sf.sheet.range(aggfunc)
    return HeatFrame(
        2,
        8,
        data=fc.sf,
        x="X",
        y="Y",
        value="v",
        aggfunc=aggfunc,
        sheet=fc.sf.sheet,
    )


def test_index(sf: HeatFrame):
    assert sf.sheet.range("H3:H6").value == [1, 2, 3, 4]


def test_columns(sf: HeatFrame):
    assert sf.sheet.range("I2:K2").value == [1, 2, 3]


@pytest.mark.parametrize(
    ("i", "value"),
    [(3, [8.5, 80.5, 152.5]), (4, [26.5, 98.5, 170.5])],
)
def test_values(sf: HeatFrame, i: int, value: int):
    assert sf.sheet.range(f"I{i}:K{i}").value == value


@pytest.fixture(scope="module")
def sf_func(fc: Agg):
    fc.sf.sheet.range("$M$13").value = "max"
    src = fc.sf.groupby(["X", "Y"]).agg("$M$13", formula=True)
    return HeatFrame(8, 8, data=src, x="X", y="Y", value="v", sheet=fc.sf.sheet)


@pytest.mark.parametrize(
    ("func", "value"),
    [("min", 0), ("max", 17), ("mean", 8.5), ("count", 18)],
)
def test_values_func(sf_func: HeatFrame, func, value):
    sf_func.sheet.range("$M$13").value = func
    assert sf_func.sheet.range("I9").value == value
