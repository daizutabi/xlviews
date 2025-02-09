import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame import WideColumn

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return WideColumn(sheet_module, 4, 2)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 4
    assert sf.column == 2
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 2
    assert sf.columns_level == 1


def test_expand(sf: SheetFrame):
    v = sf.expand().options(ndim=2).value
    assert v
    assert len(v) == 3
    assert v[0] == ["x", "y", "a", "b", *range(3), *range(4)]
    assert v[1] == ["i", "k", 1, 3, *([None] * 7)]
    assert v[2] == ["j", "l", 2, 4, *([None] * 7)]

    assert sf.cell.offset(-1, 4).value == "u"
    assert sf.cell.offset(-1, 7).value == "v"


def test_len(sf: SheetFrame):
    assert len(sf) == 2


def test_columns(sf: SheetFrame):
    assert sf.columns == ["x", "y", "a", "b", *range(3), *range(4)]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b", *range(3), *range(4)]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y"]


def test_wide_columns(sf: SheetFrame):
    assert sf.wide_columns == ["u", "v"]


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "a" in sf
    assert "u" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["x", "y", "a", "b", *range(3), *range(4)]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("a", 4),
        ("b", 5),
        (["y", "b"], [3, 5]),
        ("u", (6, 8)),
        ("v", (9, 12)),
        (("v", 0), 9),
        (("v", 3), 12),
        (["x", "a", "u", ("v", 0)], [2, 4, (6, 8), 9]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index(column) == index


@pytest.mark.parametrize("column", ["z", ("u", -1)])
def test_index_error(sf: SheetFrame, column):
    with pytest.raises(ValueError, match=".* is not in list"):
        sf.index(column)


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns[:2], df.columns)
    np.testing.assert_array_equal(df_.columns[2:], [*range(3), *range(4)])
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_.iloc[:, :2], df)
    assert df_.iloc[:, 2:].isna().all().all()
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("x", -1, "$B$4"),
        ("y", 0, "$C$5"),
        ("a", None, "$D$5:$D$6"),
        (("u", 0), -1, "$F$3:$F$4"),
        (("u", 2), 0, "$H$5"),
        (("v", 0), None, "$I$5:$I$6"),
        ("u", -1, "$F$3:$H$4"),
        ("u", 0, "$F$5:$H$5"),
        ("v", None, "$I$5:$L$6"),
    ],
)
def test_range_column(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("column", "value"),
    [("a", [1, 2]), (("u", 1), [np.nan, np.nan])],
)
def test_getitem_str(sf: SheetFrame, column, value):
    s = sf[column]
    assert isinstance(s, Series)
    assert s.name == column
    np.testing.assert_array_equal(s, value)


def test_getitem_list(sf: SheetFrame):
    df = sf[["b", ("v", 3)]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == ["b", ("v", 3)]
    x = [[3, np.nan], [4, np.nan]]
    np.testing.assert_array_equal(df, x)
