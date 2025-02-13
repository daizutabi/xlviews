import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import NoIndex

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return NoIndex(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, fc: FrameContainer):
    assert sf.row == fc.row
    assert sf.column == fc.column
    assert sf.index.nlevels == 1
    assert sf.columns.nlevels == 1
    assert sf.columns_names is None


def test_repr(sf: SheetFrame):
    assert repr(sf).endswith("!$B$2:$D$6>")


def test_str(sf: SheetFrame):
    assert str(sf).endswith("!$B$2:$D$6>")


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.headers == [None, "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == [None]


def test_contains(sf: SheetFrame):
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == [None, "a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("a", 3),
        ("b", 4),
        (["a", "b"], [3, 4]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index


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
        ("a", 0, "$C$3"),
        ("a", -1, "$C$2"),
        ("b", -1, "$D$2"),
        ("a", None, "$C$3:$C$6"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


def test_range_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="invalid offset"):
        sf.range("a", 1)  # type: ignore


def test_setitem(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    x = [10, 20, 30]
    sf["a"] = x
    np.testing.assert_array_equal(sf.data["a"], x)


def test_setitem_new_column(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    x = [10, 20, 30]
    sf["c"] = x
    assert sf.headers == [None, "a", "b", "c"]
    np.testing.assert_array_equal(sf.data["c"], x)


def test_get_address(sf: SheetFrame):
    df = sf.get_address()
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.to_list() == [0, 1, 2, 3]
    assert df.loc[0, "a"] == "$C$3"
    assert df.loc[0, "b"] == "$D$3"
    assert df.loc[1, "a"] == "$C$4"
    assert df.loc[1, "b"] == "$D$4"
    assert df.loc[2, "a"] == "$C$5"
    assert df.loc[2, "b"] == "$D$5"
    assert df.loc[3, "a"] == "$C$6"
    assert df.loc[3, "a"] == "$C$6"
