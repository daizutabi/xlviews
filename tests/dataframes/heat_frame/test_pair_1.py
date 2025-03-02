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
    return [hf for _, hf in pair(sf, values=None, columns="A", axis=1)]


def test_len(hfs: list[HeatFrame]):
    assert len(hfs) == 12


@pytest.mark.parametrize(
    ("i", "v"),
    [
        (0, 1411),
        (1, 2511),
        (2, 1711),
        (3, 2811),
        (4, None),
        (5, 3111),
        (6, 0),
        (7, 241),
        (8, 138),
        (9, 333),
        (10, None),
        (11, 402),
    ],
)
def test_value(hfs: list[HeatFrame], i: int, v):
    assert hfs[i].cell.offset(1, 1).value == v
