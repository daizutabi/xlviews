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
    return MultiColumn(sheet_module, 2, 2)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$B$2"
    assert sf.row == 2
    assert sf.column == 2
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 1
    assert sf.columns_level == 4
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


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == [("s", "t", "r", "i")]


def test_contains(sf: SheetFrame):
    assert ("s", "t", "r", "i") in sf


def test_iter(sf: SheetFrame):
    assert list(sf)[-1] == ("b", "d", 8, "y")


@pytest.mark.parametrize(
    ("column", "index"),
    [
        (("s", "t", "r", "i"), 2),
        (("a", "d", 3, "x"), 7),
        ([("b", "c", 6, "y"), ("b", "d", 8, "x")], [14, 17]),
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
        (("a", "d", 4, "y"), -1, "$J$2:$J$5"),
        (("a", "d", 3, "x"), 0, "$G$6"),
        (("b", "d", 7, "x"), None, "$O$6:$O$11"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


def test_ranges(sf: SheetFrame):
    for rng, i in zip(sf.ranges(), range(2, 18), strict=True):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


@pytest.mark.parametrize(
    ("by", "result"),
    [
        ("s", {("a",): [(3, 10)], ("b",): [(11, 18)]}),
        (
            ["s", "t"],
            {
                ("a", "c"): [(3, 6)],
                ("a", "d"): [(7, 10)],
                ("b", "c"): [(11, 14)],
                ("b", "d"): [(15, 18)],
            },
        ),
        (
            ["s", "i"],
            {
                ("a", "x"): [(3, 3), (5, 5), (7, 7), (9, 9)],
                ("a", "y"): [(4, 4), (6, 6), (8, 8), (10, 10)],
                ("b", "x"): [(11, 11), (13, 13), (15, 15), (17, 17)],
                ("b", "y"): [(12, 12), (14, 14), (16, 16), (18, 18)],
            },
        ),
        (None, {(): [(3, 18)]}),
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
        (0, ["a", "c", 1, "x", "=$C$6:$C$11"]),
        (1, ["a", "c", 1, "y", "=$D$6:$D$11"]),
        (2, ["a", "c", 2, "x", "=$E$6:$E$11"]),
        (7, ["a", "d", 4, "y", "=$J$6:$J$11"]),
        (14, ["b", "d", 8, "x", "=$Q$6:$Q$11"]),
        (15, ["b", "d", 8, "y", "=$R$6:$R$11"]),
    ],
)
def test_melt_value(df_melt: DataFrame, i, v):
    assert df_melt.iloc[i].to_list() == v
