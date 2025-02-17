import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.sheet_frame.pivot import Pivot
from xlviews.utils import iter_group_locs

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Pivot(sheet_module)


@pytest.fixture(scope="module")
def sf(fc: Pivot):
    return fc.sf


@pytest.fixture(scope="module")
def df(sf: SheetFrame):
    return sf.pivot_table("u", ["B", "Y"], ["A", "X"], "mean")


def test_index_level_values(df: DataFrame):
    v = df.index.get_level_values(0).to_list()
    assert v == [1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 3]


def test_columns_level_values(df: DataFrame):
    v = df.columns.get_level_values(0).to_list()
    assert v == [1, 1, 1, 2, 2]


def test_index_group_locs(df: DataFrame):
    v = list(iter_group_locs(df.index.get_level_values(0)))
    assert v == [(0, 3), (4, 6), (7, 10)]


def test_columns_group_locs(df: DataFrame):
    v = list(iter_group_locs(df.columns.get_level_values(0)))
    assert v == [(0, 2), (3, 4)]
