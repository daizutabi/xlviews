import numpy as np
import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.table import Table
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame import Index

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Index(sheet_module, 2, 3)


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


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("x", [[1, 5], [2, 6]]),
        ("y", [[3, 7], [4, 8]]),
    ],
)
def test_visible_data(sf: SheetFrame, table: Table, name, value):
    table.auto_filter("name", name)
    df = sf.visible_data
    assert df.index.to_list() == [name, name]
    np.testing.assert_array_equal(df, value)
