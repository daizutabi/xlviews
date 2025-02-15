import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame.base import MultiIndexColumn

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiIndexColumn(sheet_module)


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
    assert sf.cell.get_address() == "$U$13"
    assert sf.row == 13
    assert sf.column == 21
    assert sf.sheet.name == sheet_module.name
    assert sf.index.nlevels == 3
    assert sf.columns.nlevels == 2
    assert sf.columns_names is None


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["x", "y", "z"]


def test_contains(sf: SheetFrame):
    assert "x" not in sf
    assert "a1" in sf
    assert ("a1", "b1") not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("z", 23),
        (("a1", "b1"), 24),
        (["x", ("a2", "b2")], [21, 27]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index_past(column) == index


def test_index_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="'a' is not in list"):
        sf.index_past("a")


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("x", -1, "$U$13:$U$14"),
        ("y", 0, "$V$15"),
        ("z", None, "$W$15:$W$22"),
        (("a1", "b1"), -1, "$X$13:$X$14"),
        (("a1", "b2"), 0, "$Y$15"),
        (("a2", "b1"), None, "$Z$15:$Z$22"),
    ],
)
def test_range_column(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address
