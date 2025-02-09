import pytest
from xlwings import Range, Sheet

from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.mark.parametrize("index", range(1, 1000, 50))
def test_column_name(index: int):
    from xlviews.range.address import column_name_to_index, index_to_column_name

    assert column_name_to_index(index_to_column_name(index)) == index


@pytest.fixture(scope="module", params=["A1", "A1:A3", "F4:I4", "C1:E3"])
def addr(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def rng(addr, sheet_module: Sheet):
    return sheet_module.range(addr)


def test_split_book(addr, rng: Range, sheet_module: Sheet):
    from xlviews.range.address import split_book

    addr = rng.get_address(external=True)
    addr2 = rng.get_address(include_sheetname=True)
    assert split_book(addr) == (sheet_module.book.name, addr2)


def test_split_book_none(addr, rng: Range, sheet_module: Sheet):
    from xlviews.range.address import split_book

    addr = rng.get_address(include_sheetname=True)
    assert split_book(addr) == (sheet_module.book.name, addr)


def test_split_sheet(addr, rng: Range, sheet_module: Sheet):
    from xlviews.range.address import split_sheet

    addr = rng.get_address(include_sheetname=True)
    addr2 = rng.get_address()
    assert split_sheet(addr) == (rng.sheet.name, addr2)


def test_split_sheet_none(addr, rng: Range, sheet_module: Sheet):
    from xlviews.range.address import split_sheet

    addr = rng.get_address()
    assert split_sheet(addr) == (sheet_module.name, addr)


@pytest.mark.parametrize("absolute", [False, True])
def test_get_index(addr, absolute, rng: Range, sheet_module: Sheet):
    from xlviews.range.address import get_index

    addr = rng.get_address(row_absolute=absolute, column_absolute=absolute)
    x = (rng.row, rng.column, rng.last_cell.row, rng.last_cell.column)
    assert get_index(addr) == x


def test_get_index_error():
    from xlviews.range.address import get_index

    with pytest.raises(ValueError, match="Invalid address format: 11"):
        get_index("11")
