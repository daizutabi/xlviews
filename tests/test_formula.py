import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Range, Sheet

from xlviews.formula import NONCONST_VALUE


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 1, 1, 1], "b": [2, 2, 3, 3], "c": [4, 4, 4, 4]})


@pytest.fixture(scope="module")
def rng(df: DataFrame, sheet_module: Sheet):
    rng = sheet_module.range("B3")
    rng.options(DataFrame, header=1, index=False).value = df
    return rng.expand()


def test_range_address(rng: Range):
    assert rng.get_address() == "$B$3:$D$7"


@pytest.mark.parametrize("k", range(3))
def test_range_value(rng: Range, df: DataFrame, k: int):
    value = rng.options(transpose=True).value
    assert isinstance(value, list)
    assert len(value) == 3
    np.testing.assert_array_equal(df[value[k][0]], value[k][1:])


@pytest.fixture(scope="module")
def column(rng: Range):
    start = rng[0].offset(1)
    end = start.offset(3)
    return rng.sheet.range(start, end)


def test_column(column: Range):
    assert column.get_address() == "$B$4:$B$7"


@pytest.fixture(scope="module")
def const_header(rng: Range):
    end = rng[0].expand("right")
    return rng.sheet.range(rng[0], end).offset(-1)


def test_header(const_header: Range):
    assert const_header.get_address() == "$B$2:$D$2"


@pytest.mark.parametrize(("k", "value"), [(0, 1), (1, NONCONST_VALUE), (2, 4)])
def test_const(column: Range, const_header: Range, k, value):
    from xlviews.formula import const

    const_header.value = const(column, "=")
    assert const_header.value[k] == value
