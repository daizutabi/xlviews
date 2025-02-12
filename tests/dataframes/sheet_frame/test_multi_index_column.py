import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import MultiIndexColumn

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiIndexColumn(sheet_module, 5, 3)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_value_column(sf: SheetFrame):
    assert sf.cell.value is None
    v = sf.cell.offset(0, 3).expand().options(ndim=2).value
    assert len(v) == 10
    assert v[0] == ["a1", "a1", "a2", "a2"]
    assert v[1] == ["b1", "b2", "b1", "b2"]
    assert v[2] == [1, 11, 21, 31]
    assert v[-1] == [8, 18, 28, 38]


def test_value_values(sf: SheetFrame):
    v = sf.cell.offset(1, 0).expand().options(ndim=2).value
    assert len(v) == 9
    assert v[0] == ["x", "y", "z", "b1", "b2", "b1", "b2"]
    assert v[1] == [1, 1, 1, 1, 11, 21, 31]
    assert v[2] == [1, 1, 1, 2, 12, 22, 32]
    assert v[-1] == [2, 2, 1, 8, 18, 28, 38]


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$C$5"
    assert sf.row == 5
    assert sf.column == 3
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 3
    assert sf.columns_level == 2
    assert sf.columns_names is None


def test_load(sf: SheetFrame):
    sf.load(index_level=2, columns_level=2)
    assert sf.index_level == 2
    assert sf.index_columns == ["x", "y"]
    c = [(None, "z"), ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c
    sf.load(index_level=3, columns_level=2)
    assert sf.index_level == 3
    assert sf.index_columns == ["x", "y", "z"]
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c


def test_expand(sf: SheetFrame):
    v = sf.expand().options(ndim=2).value
    assert v
    assert len(v) == 10
    assert v[0] == [None, None, None, "a1", "a1", "a2", "a2"]
    assert v[1] == ["x", "y", "z", "b1", "b2", "b1", "b2"]
    assert v[-1] == [2, 2, 1, 8, 18, 28, 38]


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_columns(sf: SheetFrame):
    i = "x", "y", "z"
    c = ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")
    assert sf.columns == [*i, *c]


def test_value_columns(sf: SheetFrame):
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y", "z"]


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "z" in sf
    assert ("a1", "b1") in sf
    assert "a1" not in sf


def test_iter(sf: SheetFrame):
    i = "x", "y", "z"
    c = ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")
    assert list(sf) == [*i, *c]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("z", 5),
        (("a1", "b1"), 6),
        (["x", ("a2", "b2")], [3, 9]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index(column) == index


def test_index_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="'a' is not in list"):
        sf.index("a")


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("x", -1, "$C$5:$C$6"),
        ("y", 0, "$D$7"),
        ("z", None, "$E$7:$E$14"),
        (("a1", "b1"), -1, "$F$5:$F$6"),
        (("a1", "b2"), 0, "$G$7"),
        (("a2", "b1"), None, "$H$7:$H$14"),
    ],
)
def test_range_column(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("by", "one", "two"),
    [
        ("x", [(7, 10)], [(11, 14)]),
        ("y", [(7, 8), (11, 12)], [(9, 10), (13, 14)]),
        ("z", [(7, 9), (12, 14)], [(10, 11)]),
    ],
)
def test_groupby(sf: SheetFrame, by, one, two):
    g = groupby(sf, by)
    assert len(g) == 2
    assert g[(1,)] == one
    assert g[(2,)] == two


@pytest.mark.parametrize(
    ("by", "v11", "v12", "v21", "v22"),
    [
        (["x", "y"], [(7, 8)], [(9, 10)], [(11, 12)], [(13, 14)]),
        (["x", "z"], [(7, 9)], [(10, 10)], [(12, 14)], [(11, 11)]),
        (["y", "z"], [(7, 8), (12, 12)], [(11, 11)], [(9, 9), (13, 14)], [(10, 10)]),
    ],
)
def test_groupby_list(sf: SheetFrame, by, v11, v12, v21, v22):
    g = groupby(sf, by)
    assert len(g) == 4
    assert g[(1, 1)] == v11
    assert g[(1, 2)] == v12
    assert g[(2, 1)] == v21
    assert g[(2, 2)] == v22
