import string

import pytest
from pandas import DataFrame
from xlwings import Sheet

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


def test_len(sf: SheetFrame):
    assert len(sf) == 6


def test_index_names(sf: SheetFrame):
    assert sf.index.names == [None]


def test_columns_names(sf: SheetFrame):
    assert sf.columns.names == ["s", "t", "r", "i"]


def test_iter(sf: SheetFrame):
    assert list(sf)[-1] == ("b", "d", 8, "y")


@pytest.mark.parametrize(
    ("column", "index"),
    [
        (("a", "d", 3, "x"), 16),
        ([("b", "c", 6, "y"), ("b", "d", 8, "x")], [23, 26]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index


@pytest.mark.parametrize("column", ["s", "t", "r", "i"])
def test_index_row(sf: SheetFrame, column):
    r = sf.index_past(column)
    assert sf.sheet.range(r, sf.column).value == column


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
