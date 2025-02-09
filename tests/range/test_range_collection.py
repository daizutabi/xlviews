import pytest
from xlwings import Sheet

from xlviews.range.range_collection import RangeCollection
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.mark.parametrize(
    ("row", "n", "address"),
    [
        ([(4, 5), (10, 14)], 2, "E4:E5,E10:E14"),
        ([(5, 5), (7, 8), (10, 11)], 3, "E5,E7:E8,E10:E11"),
        ([1, 5], 2, "E1,E5"),
        (4, 1, "E4"),
    ],
)
def test_range_collection_row(row, n, address, sheet_module: Sheet):
    rc = RangeCollection(row, 5, sheet_module)
    assert len(rc) == n
    a = rc.get_address(row_absolute=False, column_absolute=False)
    assert a == address


@pytest.mark.parametrize(
    ("column", "n", "address"),
    [
        ([(2, 2)], 1, "$B$5"),
        ([(4, 5), (10, 14)], 2, "$D$5:$E$5,$J$5:$N$5"),
        ([(5, 5), (7, 8), (10, 11)], 3, "$E$5,$G$5:$H$5,$J$5:$K$5"),
        ([1, 5], 2, "$A$5,$E$5"),
        (4, 1, "$D$5"),
    ],
)
def test_range_collection_column(column, n, address, sheet_module: Sheet):
    rc = RangeCollection(5, column, sheet_module)
    assert len(rc) == n
    assert rc.get_address() == address
    assert rc.api.Address == address


@pytest.mark.parametrize(
    ("ranges", "n"),
    [
        (["A1:B3"], 1),
        (["A2:A4", "A5:A8"], 2),
        (["A2:A4,A5:A8"], 2),
        (["A2:A4,A5:A8", "C4:C7,D10:D12"], 4),
    ],
)
def test_range_collection_from_str(ranges, n, sheet_module: Sheet):
    rc = RangeCollection.from_iterable(ranges)
    assert len(rc) == n
    a = rc.get_address(row_absolute=False, column_absolute=False)
    assert a == ",".join(ranges)


@pytest.mark.parametrize(
    ("ranges", "n"),
    [(["$A$1:$B$3"], 1), (["$A$2:$A$4", "$A$5:$A$8"], 2)],
)
def test_range_collection_from_range(ranges, n, sheet_module: Sheet):
    ranges = [sheet_module.range(r) for r in ranges]
    rc = RangeCollection.from_iterable(ranges)
    assert len(rc) == n
    a = rc.get_address()
    assert a == ",".join(r.get_address() for r in ranges)


def test_range_collection_from_range_one(sheet_module: Sheet):
    rng = sheet_module.range("A1:B2")
    rc = RangeCollection.from_iterable(rng)
    assert len(rc) == 4
    a = rc.get_address()
    assert a == "$A$1,$B$1,$A$2,$B$2"


def test_range_collection_from_range_list(sheet_module: Sheet):
    rng = sheet_module.range("A1:B2")
    rc = RangeCollection.from_iterable([rng])
    assert len(rc) == 1
    a = rc.get_address()
    assert a == "$A$1:$B$2"


def test_range_collection_iter(sheet_module: Sheet):
    rc = RangeCollection([(2, 5), (10, 12)], 1, sheet_module)
    for rng, row in zip(rc, [2, 10], strict=True):
        assert rng.row == row


def test_range_collection_repr(sheet_module: Sheet):
    rc = RangeCollection([(2, 5), (8, 10)], 5, sheet_module)
    assert repr(rc) == "<RangeCollection $E$2:$E$5,$E$8:$E$10>"


def test_iter_ranges_error(sheet_module: Sheet):
    with pytest.raises(TypeError):
        RangeCollection((1, 2), (3, 4), sheet_module)
