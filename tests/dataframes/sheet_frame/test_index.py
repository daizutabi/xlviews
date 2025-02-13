import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import Index

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Index(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame):
    assert sf.row == 8
    assert sf.column == 2
    assert sf.index_level == 1
    assert sf.columns_level == 1
    assert sf.columns_names is None


def test_load(sf: SheetFrame):
    sf.load(index_level=0)
    assert sf.index_level == 0
    assert sf.value_columns == ["name", "a", "b"]
    sf.load(index_level=1)
    assert sf.index_level == 1
    assert sf.value_columns == ["a", "b"]


def test_expand(sf: SheetFrame):
    v = [["name", "a", "b"], ["x", 1, 5], ["x", 2, 6], ["y", 3, 7], ["y", 4, 8]]
    assert sf.expand().options(ndim=2).value == v


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.headers == ["name", "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["name"]


def test_contains(sf: SheetFrame):
    assert "name" in sf
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["name", "a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("name", 2),
        ("a", 3),
        ("b", 4),
        (["name", "b"], [2, 4]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index
    assert sf.column_index(column) == index


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("b", 0, "$D$9"),
        ("a", -1, "$C$8"),
        ("b", -1, "$D$8"),
        ("name", -1, "$B$8"),
        ("a", None, "$C$9:$C$12"),
    ],
)
def test_range(sf: SheetFrame, column: str, offset, address):
    assert sf.range(column, offset).get_address() == address
    assert sf.column_range(column, offset).get_address() == address


def test_get_address(sf: SheetFrame):
    df = sf.get_address(row_absolute=False, column_absolute=False, formula=True)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.name == "name"
    assert df.index.to_list() == ["x", "x", "y", "y"]
    assert df.to_numpy()[0, 0] == "=C9"
