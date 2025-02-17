import numpy as np
import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.base import MultiIndex, MultiIndexParent

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return MultiIndexParent(sheet_module)


@pytest.fixture(scope="module")
def sf(fc_parent: MultiIndexParent):
    fc = MultiIndex(fc_parent.sf)
    return fc.sf


def test_index(sf: HeatFrame):
    x = sf.sheet.range("V3:V26").value
    assert x
    y = np.array([None] * 24)
    y[::6] = range(1, 5)
    np.testing.assert_array_equal(x, y)


def test_index_from_sf(sf: HeatFrame):
    x = np.repeat(list(range(1, 5)), 6)
    np.testing.assert_array_equal(sf.index, x)


def test_columns(sf: HeatFrame):
    x = sf.sheet.range("W2:AH2").value
    assert x
    y = np.array([None] * 12)
    y[::4] = range(1, 4)
    np.testing.assert_array_equal(x, y)


def test_columns_from_sf(sf: HeatFrame):
    x = np.repeat(list(range(1, 4)), 4)
    np.testing.assert_array_equal(sf.columns, x)


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, [0, 5, None, 13, 72]),
        (4, [1, None, 9, 14, 73]),
        (5, [None, 6, 10, 15, None]),
    ],
)
def test_values(sf: HeatFrame, i: int, value: int):
    assert sf.sheet.range(f"W{i}:AA{i}").value == value
