import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.facet_agg import facet
from xlviews.testing.sheet_frame.pivot import Pivot

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Pivot(sheet_module)


@pytest.fixture(scope="module")
def df(fc: Pivot):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: Pivot):
    return fc.sf


@pytest.fixture(scope="module")
def hfs(sf: SheetFrame):
    return [hf for _, hf in facet(sf)]


@pytest.mark.parametrize("i", [0, 1, 2])
@pytest.mark.parametrize("j", [0, 1])
def test_value(hfs: list[HeatFrame], df: DataFrame, i: int, j: int):
    x = hfs[i].cell.offset(1, j + 4).value
    df = df.reset_index()
    a = df[(df["B"] == i + 1) & (df["Y"] == 1) & (df["A"] == 2) & (df["X"] == j + 2)]
    y = a["u"].mean()
    assert x == y
