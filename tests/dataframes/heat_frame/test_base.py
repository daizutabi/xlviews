import numpy as np
import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.base import Base, BaseParent

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return BaseParent(sheet_module)


@pytest.fixture(scope="module")
def sf(fc_parent: BaseParent):
    fc = Base(fc_parent.sf)
    return fc.sf


def test_index(sf: HeatFrame):
    assert sf.sheet.range("G3:G8").value == [1, 2, 3, 4, 5, 6]


def test_index_from_sf(sf: HeatFrame):
    assert sf.index.to_list() == [1, 2, 3, 4, 5, 6]


def test_columns(sf: HeatFrame):
    assert sf.sheet.range("H2:K2").value == [1, 2, 3, 4]


def test_columns_from_sf(sf: HeatFrame):
    assert sf.columns.to_list() == [1, 2, 3, 4]


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, [0, 6, None, 18]),
        (4, [1, None, 13, 19]),
        (5, [None, 8, 14, 20]),
        (6, [3, 9, 15, None]),
        (7, [4, 10, None, 22]),
        (8, [5, None, 17, 23]),
    ],
)
def test_values(sf: HeatFrame, i: int, value: int):
    assert sf.sheet.range(f"H{i}:K{i}").value == value


def test_label(sf: HeatFrame):
    assert sf.sheet.range("M2").value == "v"


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, 23),
        (4, 23 * 4 / 5),
        (5, 23 * 3 / 5),
        (6, 23 * 2 / 5),
        (7, 23 / 5),
        (8, 0),
    ],
)
def test_colorbar(sf: HeatFrame, i: int, value: int):
    v = sf.sheet.range(f"M{i}").value
    assert isinstance(v, float)
    np.testing.assert_allclose(v, value)


def test_number_format(sf: HeatFrame):
    sf.number_format("0.00")
    assert sf.sheet.range("H4").api.NumberFormatLocal == "0.00"
    assert sf.sheet.range("K8").api.NumberFormatLocal == "0.00"
