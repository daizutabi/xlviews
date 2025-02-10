import numpy as np
import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_excel_installed
from xlviews.testing.heat_frame.base import MultiIndex

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiIndex(sheet_module)


@pytest.fixture(scope="module")
def sf(fc: MultiIndex):
    x = ["X", "x"]
    y = ["Y", "y"]
    return HeatFrame(2, 20, data=fc.sf, x=x, y=y, value="v", sheet=fc.sf.sheet)


def test_index(sf: HeatFrame):
    x = sf.sheet.range("T3:T26").value
    assert x
    y = np.array([None] * 24)
    y[::6] = range(1, 5)
    np.testing.assert_array_equal(x, y)


def test_index_from_df(sf: HeatFrame):
    x = np.repeat(list(range(1, 5)), 6)
    np.testing.assert_array_equal(sf.df.index, x)


def test_columns(sf: HeatFrame):
    x = sf.sheet.range("U2:AF2").value
    assert x
    y = np.array([None] * 12)
    y[::4] = range(1, 4)
    np.testing.assert_array_equal(x, y)


def test_columns_from_df(sf: HeatFrame):
    x = np.repeat(list(range(1, 4)), 4)
    np.testing.assert_array_equal(sf.df.columns, x)


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, [0, 5, None, 13, 72]),
        (4, [1, None, 9, 14, 73]),
        (5, [None, 6, 10, 15, None]),
    ],
)
def test_values(sf: HeatFrame, i: int, value: int):
    assert sf.sheet.range(f"U{i}:Y{i}").value == value


def test_vmin(sf: HeatFrame):
    assert sf.vmin.get_address() == "$AH$26"


def test_vmax(sf: HeatFrame):
    assert sf.vmax.get_address() == "$AH$3"


def test_label(sf: HeatFrame):
    assert sf.label.get_address() == "$AH$2"


def test_label_value(sf: HeatFrame):
    assert sf.label.value == "v"
