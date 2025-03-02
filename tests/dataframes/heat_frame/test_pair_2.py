import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.pair import pair
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
    return [hf for _, hf in pair(sf, values="u", columns=None)]


def test_len(hfs: list[HeatFrame]):
    assert len(hfs) == 3


@pytest.mark.parametrize(
    ("i", "v"),
    [(0, 2614), (1, 2914), (2, 3214)],
)
def test_value(hfs: list[HeatFrame], i: int, v):
    assert hfs[i].cell.offset(1, 15).value == v
