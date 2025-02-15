import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import Index

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


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
    assert sf.columns_names is None


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns.to_list() == ["a", "b"]


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["name"]


def test_contains(sf: SheetFrame):
    assert "name" not in sf
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("name", 2),
        ("a", 3),
        ("b", 4),
        (["name", "b"], [2, 4]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.column_index(column) == index


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
def test_column_range(sf: SheetFrame, column: str, offset, address):
    assert sf.column_range(column, offset).get_address() == address


def test_get_address(sf: SheetFrame):
    df = sf.get_address(row_absolute=False, column_absolute=False, formula=True)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.name == "name"
    assert df.index.to_list() == ["x", "x", "y", "y"]
    assert df.to_numpy()[0, 0] == "=C9"
