import pytest
from xlwings import Sheet

from xlviews.range.range_collection import RangeCollection
from xlviews.utils import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def test_range_value_int(sheet: Sheet):
    sheet.range(1, 1).value = 10
    x = sheet.range(1, 1).value
    assert not isinstance(x, int)
    assert isinstance(x, float)
    assert x == 10


def test_range_value_str(sheet: Sheet):
    sheet.range(1, 1).value = "abc"
    x = sheet.range(1, 1).value
    assert isinstance(x, str)
    assert x == "abc"


@pytest.mark.parametrize(
    ("ranges", "n"),
    [
        (["A1:B3"], 1),
        (["A2:A4", "A5:A8"], 2),
        (["A2:A4,A5:A8"], 1),
        (["A2:A4,A5:A8", "C4:C7,D10:D12"], 2),
    ],
)
def test_range_collection_from_str(sheet_module: Sheet, ranges, n):
    rc = RangeCollection(ranges)
    assert len(rc) == n
    a = rc.get_address(row_absolute=False, column_absolute=False)
    assert a == ",".join(ranges)


@pytest.mark.parametrize(
    ("row", "n", "address"),
    [
        ([(4, 5), (10, 14)], 2, "E4:E5,E10:E14"),
        ([(5, 5), (7, 8), (10, 11)], 3, "E5,E7:E8,E10:E11"),
    ],
)
def test_range_collection_from_index_row(sheet_module: Sheet, row, n, address):
    rc = RangeCollection.from_index(sheet_module, row, 5)
    assert len(rc) == n
    a = rc.get_address(row_absolute=False, column_absolute=False)
    assert a == address


@pytest.mark.parametrize(
    ("column", "n", "address"),
    [
        ([(2, 2)], 1, "$B$5"),
        ([(4, 5), (10, 14)], 2, "$D$5:$E$5,$J$5:$N$5"),
        ([(5, 5), (7, 8), (10, 11)], 3, "$E$5,$G$5:$H$5,$J$5:$K$5"),
    ],
)
def test_range_collection_from_index_column(sheet_module: Sheet, column, n, address):
    rc = RangeCollection.from_index(sheet_module, 5, column)
    assert len(rc) == n
    assert rc.get_address() == address
    assert rc.api.Address == address


def test_range_collection_iter(sheet_module: Sheet):
    rc = RangeCollection.from_index(sheet_module, [(2, 5), (10, 12)], 1)
    for rng, row in zip(rc, [2, 10], strict=True):
        assert rng.row == row


def test_range_collection_repr(sheet_module: Sheet):
    rc = RangeCollection.from_index(sheet_module, [(2, 5), (8, 10)], 5)
    assert repr(rc) == "<RangeCollection $E$2:$E$5,$E$8:$E$10>"
