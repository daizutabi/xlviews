import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.table import Table
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


@pytest.fixture
def table(sf: SheetFrame):
    yield sf.as_table()
    sf.unlist()


@pytest.mark.parametrize("value", ["x", "y"])
def test_table(table: Table, value):
    table.auto_filter("name", value)
    header = table.const_header.value
    assert isinstance(header, list)
    assert header[0] == value
