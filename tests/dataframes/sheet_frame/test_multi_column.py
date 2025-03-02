import string

import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import MultiColumn

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


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


def test_len(sf: SheetFrame):
    assert len(sf) == 6


def test_index_names(sf: SheetFrame):
    assert sf.index.names == [None]


def test_columns_names(sf: SheetFrame):
    assert sf.columns.names == ["s", "t", "r", "i"]


def test_iter(sf: SheetFrame):
    assert list(sf)[-1] == ("b", "d", 8, "y")


def test_value(sf: SheetFrame, df: DataFrame):
    df_sf = sf.value
    assert df_sf.equals(df.astype(float))
    assert df_sf.index.equals(df.index)
    assert df_sf.columns.equals(df.columns)


@pytest.mark.parametrize(
    ("columns", "indexer"),
    [
        ({"s": "a", "t": "c"}, [12, 13, 14, 15]),
        ({"r": 4, "i": "x"}, [18]),
        ({"t": "c", "i": "x"}, [12, 14, 20, 22]),
    ],
)
def test_get_indexer(sf: SheetFrame, columns, indexer):
    assert all(sf.get_indexer(columns) == indexer)
    assert all(sf.get_indexer(**columns) == indexer)


def test_iter_ranges(sf: SheetFrame):
    for rng, i in zip(sf.iter_ranges(), range(11, 26), strict=False):
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


def test_agg(sf: SheetFrame):
    df_sf = sf.agg()
    df_melt = sf.melt()
    assert isinstance(df_sf, DataFrame)
    assert df_sf.equals(df_melt)
    assert df_sf.index.equals(df_melt.index)
    assert df_sf.columns.equals(df_melt.columns)


def test_number_format(sf: SheetFrame):
    sf.number_format("0.000", s="a", r=3, autofit=True)
    i = sf.get_indexer({"s": "a", "r": 3})
    assert sf.sheet.range(6, i[0]).number_format == "0.000"
