import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import WideColumn

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
    assert len(v) == 6
    assert v[0] == ["x", "y", "a", "b", *range(3), *range(4)]
    assert v[1] == ["i", "k", 0, 10, *([None] * 7)]
    assert v[2] == ["i", "l", 1, 11, *([None] * 7)]

    assert sf.cell.offset(-1, 4).value == "u"
    assert sf.cell.offset(-1, 7).value == "v"


def test_len(sf: SheetFrame):
    assert len(sf) == 5


def test_columns(sf: SheetFrame):
    assert sf.headers == ["x", "y", "a", "b", *range(3), *range(4)]


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
    assert sf.index_past(column) == index


@pytest.mark.parametrize("column", ["z", ("u", -1)])
def test_index_error(sf: SheetFrame, column):
    with pytest.raises(ValueError, match=".* is not in list"):
        sf.index_past(column)


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
        ("a", None, "$D$5:$D$9"),
        (("u", 0), -1, "$F$3:$F$4"),
        (("u", 2), 0, "$H$5"),
        (("v", 0), None, "$I$5:$I$9"),
        ("u", -1, "$F$3:$H$4"),
        ("u", 0, "$F$5:$H$5"),
        ("v", None, "$I$5:$L$9"),
    ],
)
def test_range_column(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("by", "v1", "v2"),
    [
        ("x", [(5, 6), (9, 9)], [(7, 8)]),
        ("y", [(5, 5), (7, 7), (9, 9)], [(6, 6), (8, 8)]),
    ],
)
def test_groupby(sf: SheetFrame, by, v1, v2):
    g = groupby(sf, by)
    assert len(g) == 2
    keys = list(g.keys())
    assert g[keys[0]] == v1
    assert g[keys[1]] == v2


def test_groupby_list(sf: SheetFrame):
    g = groupby(sf, ["x", "y"])
    assert len(g) == 4
    assert g[("i", "k")] == [(5, 5), (9, 9)]
    assert g[("i", "l")] == [(6, 6)]
    assert g[("j", "k")] == [(7, 7)]
    assert g[("j", "l")] == [(8, 8)]
