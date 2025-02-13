import string

import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import MultiColumn

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiColumn(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$K$2"
    assert sf.row == 2
    assert sf.column == 11
    assert sf.sheet.name == sheet_module.name
    assert sf.index.nlevels == 1
    assert sf.columns.nlevels == 4
    assert sf.columns_names == ["s", "t", "r", "i"]


def test_expand(sf: SheetFrame):
    v = sf.expand().options(ndim=2).value
    assert v
    assert len(v) == 10
    assert v[2][:5] == ["r", 1, 1, 2, 2]
    assert v[3][:5] == ["i", "x", "y", "x", "y"]
    assert v[-1][:5] == [5, 5, 11, 17, 23]


def test_len(sf: SheetFrame):
    assert len(sf) == 6


def test_columns(sf: SheetFrame):
    columns = sf.headers
    assert columns[0] == ("s", "t", "r", "i")
    assert columns[1] == ("a", "c", 1, "x")
    assert columns[-1] == ("b", "d", 8, "y")


def test_value_columns(sf: SheetFrame):
    columns = sf.value_columns
    assert columns[0] == ("a", "c", 1, "x")
    assert columns[-1] == ("b", "d", 8, "y")


def test_index_names(sf: SheetFrame):
    assert sf.index.names == [None]


def test_columns_names(sf: SheetFrame):
    assert sf.columns.names == ["s", "t", "r", "i"]


def test_contains(sf: SheetFrame):
    assert ("s", "t", "r", "i") in sf


def test_iter(sf: SheetFrame):
    assert list(sf)[-1] == ("b", "d", 8, "y")


@pytest.mark.parametrize(
    ("column", "index"),
    [
        (("s", "t", "r", "i"), 11),
        (("a", "d", 3, "x"), 16),
        ([("b", "c", 6, "y"), ("b", "d", 8, "x")], [23, 26]),
        ("s", 2),
        (["t", "i"], [3, 5]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index


@pytest.mark.parametrize("column", ["s", "t", "r", "i"])
def test_index_row(sf: SheetFrame, column):
    r = sf.index_past(column)
    assert sf.sheet.range(r, sf.column).value == column


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
        (("a", "d", 4, "y"), -1, "$S$2:$S$5"),
        (("a", "d", 3, "x"), 0, "$P$6"),
        (("b", "d", 7, "x"), None, "$X$6:$X$11"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


def test_ranges(sf: SheetFrame):
    for rng, i in zip(sf.ranges(), range(11, 26), strict=False):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


@pytest.mark.parametrize(
    ("by", "result"),
    [
        ("s", {("a",): [(12, 19)], ("b",): [(20, 27)]}),
        (
            ["s", "t"],
            {
                ("a", "c"): [(12, 15)],
                ("a", "d"): [(16, 19)],
                ("b", "c"): [(20, 23)],
                ("b", "d"): [(24, 27)],
            },
        ),
        (
            ["s", "i"],
            {
                ("a", "x"): [(12, 12), (14, 14), (16, 16), (18, 18)],
                ("a", "y"): [(13, 13), (15, 15), (17, 17), (19, 19)],
                ("b", "x"): [(20, 20), (22, 22), (24, 24), (26, 26)],
                ("b", "y"): [(21, 21), (23, 23), (25, 25), (27, 27)],
            },
        ),
        (None, {(): [(12, 27)]}),
    ],
)
def test_groupby(sf: SheetFrame, by, result):
    g = groupby(sf, by)
    assert len(g) == len(result)
    for k, v in g.items():
        assert result[k] == v


@pytest.fixture(scope="module")
def df_melt(sf: SheetFrame):
    return sf.melt(formula=True, value_name="v")


def test_melt_len(df_melt: DataFrame):
    assert len(df_melt) == 16


def test_melt_columns(df_melt: DataFrame):
    assert df_melt.columns.to_list() == ["s", "t", "r", "i", "v"]


@pytest.mark.parametrize(
    ("i", "v"),
    [
        (0, ["a", "c", 1, "x", "=$L$6:$L$11"]),
        (1, ["a", "c", 1, "y", "=$M$6:$M$11"]),
        (2, ["a", "c", 2, "x", "=$N$6:$N$11"]),
        (7, ["a", "d", 4, "y", "=$S$6:$S$11"]),
        (14, ["b", "d", 8, "x", "=$Z$6:$Z$11"]),
        (15, ["b", "d", 8, "y", "=$AA$6:$AA$11"]),
    ],
)
def test_melt_value(df_melt: DataFrame, i, v):
    assert df_melt.iloc[i].to_list() == v
