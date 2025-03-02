import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import WideColumn

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return WideColumn(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 3
    assert sf.column == 29
    assert sf.sheet.name == sheet_module.name
    assert sf.index.nlevels == 2
    assert sf.columns.nlevels == 1


def test_len(sf: SheetFrame):
    assert len(sf) == 5


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["x", "y"]


def test_contains(sf: SheetFrame):
    assert "x" not in sf
    assert "a" in sf
    assert "u" in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b", "u", "v"]


def test_get_number_format(sf: SheetFrame):
    assert sf.number_format(u="0.00")
    assert sf.get_number_format("u") == "0.00"
