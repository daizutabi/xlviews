import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import NoIndex

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return NoIndex(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, fc: FrameContainer):
    assert sf.row == fc.row
    assert sf.column == fc.column
    assert sf.index.nlevels == 1
    assert sf.columns.nlevels == 1


def test_repr(sf: SheetFrame):
    assert repr(sf).endswith("!$B$2:$D$6>")


def test_str(sf: SheetFrame):
    assert str(sf).endswith("!$B$2:$D$6>")


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_index_names(sf: SheetFrame):
    assert sf.index.names == [None]


def test_contains(sf: SheetFrame):
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("a", 0, "$C$3"),
        ("a", -1, "$C$2"),
        ("b", -1, "$D$2"),
        ("a", None, "$C$3:$C$6"),
    ],
)
def test_get_range(sf: SheetFrame, column, offset, address):
    assert sf.get_range(column, offset).get_address() == address


def test_get_address(sf: SheetFrame):
    df = sf.get_address()
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.to_list() == [0, 1, 2, 3]
    assert df.loc[0, "a"] == "$C$3"
    assert df.loc[0, "b"] == "$D$3"
    assert df.loc[1, "a"] == "$C$4"
    assert df.loc[1, "b"] == "$D$4"
    assert df.loc[2, "a"] == "$C$5"
    assert df.loc[2, "b"] == "$D$5"
    assert df.loc[3, "a"] == "$C$6"
    assert df.loc[3, "a"] == "$C$6"
