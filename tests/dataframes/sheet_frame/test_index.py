import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import Index

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


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
    assert sf.index.nlevels == 1
    assert sf.columns.nlevels == 1


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns.to_list() == ["a", "b"]


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["name"]


@pytest.mark.parametrize(("x", "b"), [("name", False), ("a", True)])
def test_contains(sf: SheetFrame, x, b):
    assert (x in sf) is b


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


def test_value(sf: SheetFrame, df: DataFrame):
    df_sf = sf.value
    assert df_sf.equals(df.astype(float))
    assert df_sf.index.equals(df.index)
    assert df_sf.columns.equals(df.columns)


@pytest.mark.parametrize(("column", "index"), [("name", 2), ("a", 3), ("b", 4)])
def test_loc(sf: SheetFrame, column, index):
    assert sf.get_loc(column) == index


@pytest.mark.parametrize(("column", "index"), [(["name", "b"], [2, 4])])
def test_indexer(sf: SheetFrame, column, index):
    assert sf.get_indexer(column) == index


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
def test_get_range(sf: SheetFrame, column: str, offset, address):
    assert sf.get_range(column, offset).get_address() == address


def test_get_address(sf: SheetFrame):
    df = sf.get_address(row_absolute=False, column_absolute=False, formula=True)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.name == "name"
    assert df.index.to_list() == ["x", "x", "y", "y"]
    assert df.to_numpy()[0, 0] == "=C9"


def test_get_adjacent_cell(sf: SheetFrame):
    assert sf.get_adjacent_cell(0).get_address() == "$F$8"
