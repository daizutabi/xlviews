import pytest
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.agg import AggParent, AggRange, AggStr

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return AggParent(sheet_module)


@pytest.fixture(scope="module")
def sf_str(fc_parent: AggParent):
    return AggStr(fc_parent.sf).sf


def test_index(sf_str: HeatFrame):
    assert sf_str.sheet.range("I3:I6").value == [1, 2, 3, 4]


def test_columns(sf_str: HeatFrame):
    assert sf_str.sheet.range("J2:L2").value == [1, 2, 3]


@pytest.mark.parametrize(
    ("i", "value"),
    [(3, [8.5, 80.5, 152.5]), (4, [26.5, 98.5, 170.5])],
)
def test_values(sf_str: HeatFrame, i: int, value: int):
    assert sf_str.sheet.range(f"J{i}:L{i}").value == value


@pytest.fixture(scope="module")
def sf_range(fc_parent: AggParent):
    return AggRange(fc_parent.sf).sf


@pytest.mark.parametrize(
    ("func", "value"),
    [("min", 0), ("max", 17), ("mean", 8.5), ("count", 18)],
)
def test_values_func(sf_range: HeatFrame, func, value):
    sf_range.sheet.range("$N$13").value = func
    assert sf_range.sheet.range("J9").value == value
