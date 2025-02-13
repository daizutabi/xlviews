import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import MultiIndex

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiIndex(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 2
    assert sf.column == 6
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 2
    assert sf.columns_level == 1
    assert sf.columns_names is None


def test_load(sf: SheetFrame):
    sf.load(index_level=0)
    assert sf.index_level == 0
    assert sf.index_columns == []
    assert sf.value_columns == ["x", "y", "a", "b"]
    sf.load(index_level=2)
    assert sf.index_level == 2
    assert sf.index_columns == ["x", "y"]
    assert sf.value_columns == ["a", "b"]


def test_expand(sf: SheetFrame):
    v = sf.expand().options(ndim=2).value
    assert v
    assert len(v) == 9
    assert v[0] == ["x", "y", "a", "b"]
    assert v[1] == [1, 1, 1, 11]
    assert v[2] == [1, 1, 2, 12]
    assert v[3] == [1, 2, 3, 13]
    assert v[4] == [1, 2, 4, 14]
    assert v[5] == [2, 1, 5, 15]
    assert v[6] == [2, 1, 6, 16]
    assert v[7] == [2, 2, 7, 17]
    assert v[8] == [2, 2, 8, 18]


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_columns(sf: SheetFrame):
    assert sf.headers == ["x", "y", "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y"]


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "a" in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["x", "y", "a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("a", 8),
        ("b", 9),
        (["y", "b"], [7, 9]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index


def test_index_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="'z' is not in list"):
        sf.index_past("z")


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
        ("x", -1, "$F$2"),
        ("y", 0, "$G$3"),
        ("a", -1, "$H$2"),
        ("b", 0, "$I$3"),
        ("y", None, "$G$3:$G$10"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("by", "v1", "v2"),
    [
        ("x", [(3, 6)], [(7, 10)]),
        ("y", [(3, 4), (7, 8)], [(5, 6), (9, 10)]),
    ],
)
def test_groupby(sf: SheetFrame, by, v1, v2):
    g = groupby(sf, by)
    assert len(g) == 2
    assert g[(1,)] == v1
    assert g[(2,)] == v2


def test_groupby_list(sf: SheetFrame):
    g = groupby(sf, ["x", "y"])
    assert len(g) == 4
    assert g[(1, 1)] == [(3, 4)]
    assert g[(1, 2)] == [(5, 6)]
    assert g[(2, 1)] == [(7, 8)]
    assert g[(2, 2)] == [(9, 10)]


def test_get_address(sf: SheetFrame):
    df = sf.get_address(row_absolute=False, column_absolute=False, formula=True)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.names == ["x", "y"]
    assert df.to_numpy()[0, 0] == "=H3"
