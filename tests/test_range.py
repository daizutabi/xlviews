import pytest
from xlwings import Sheet

from xlviews.range import RangeCollection


def test_reference_str(sheet_module: Sheet):
    from xlviews.range import reference

    assert reference("x", sheet_module) == "x"


def test_reference_range(sheet_module: Sheet):
    from xlviews.range import reference

    cell = sheet_module.range(4, 5)

    ref = reference(cell)
    assert ref == f"={sheet_module.name}!$E$4"


def test_reference_tuple(sheet_module: Sheet):
    from xlviews.range import reference

    ref = reference((4, 5), sheet_module)
    assert ref == f"={sheet_module.name}!$E$4"


def test_reference_error(sheet_module: Sheet):
    from xlviews.range import reference

    with pytest.raises(ValueError, match="sheet is required when `cell` is a tuple"):
        reference((4, 5))


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


def test_multirange_int_int(sheet_module: Sheet):
    from xlviews.range import multirange

    assert multirange(sheet_module, 3, 5).get_address() == "$E$3"


def test_multirange_error(sheet_module: Sheet):
    from xlviews.range import multirange

    with pytest.raises(TypeError):
        multirange(sheet_module, [3, 3], [5, 5])


@pytest.mark.parametrize(
    ("index", "n", "rng"),
    [
        ([3], 1, "$E$3"),
        ([(3, 5)], 3, "$E$3:$E$5"),
        ([(3, 5), 7], 4, "$E$3:$E$5,$E$7"),
        ([(3, 5), (7, 10)], 7, "$E$3:$E$5,$E$7:$E$10"),
    ],
)
def test_multirange_row(sheet_module: Sheet, index, n, rng):
    from xlviews.range import multirange

    x = multirange(sheet_module, index, 5)
    assert len(x) == n
    assert x.get_address() == rng


@pytest.mark.parametrize(
    ("index", "n", "rng"),
    [
        ([3], 1, "$C$10"),
        ([(3, 5)], 3, "$C$10:$E$10"),
        ([(3, 5), 7], 4, "$C$10:$E$10,$G$10"),
        ([(3, 5), (7, 10)], 7, "$C$10:$E$10,$G$10:$J$10"),
    ],
)
def test_multirange_column(sheet_module: Sheet, index, n, rng):
    from xlviews.range import multirange

    x = multirange(sheet_module, 10, index)
    assert len(x) == n
    assert x.get_address() == rng


@pytest.mark.parametrize(
    ("ranges", "n"),
    [
        (["A1:B3"], 6),
        (["A2:A4", "A5:A8"], 7),
        (["A2:A4,A5:A8"], 7),
        (["A2:A4,A5:A8", "C4:C7,D10:D12"], 14),
    ],
)
def test_range_collection_from_str(sheet_module: Sheet, ranges, n):
    rc = RangeCollection(ranges)
    assert len(rc) == n
    a = rc.get_address(row_absolute=False, column_absolute=False)
    assert a == ",".join(ranges)
    assert rc.first().sheet.name == sheet_module.name


@pytest.mark.parametrize(
    ("row", "n", "address"),
    [
        ([(4, 5), (10, 14)], 7, "E4:E5,E10:E14"),
        ([(5, 5), (7, 8), (10, 11)], 5, "E5,E7:E8,E10:E11"),
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
        ([(4, 5), (10, 14)], 7, "$D$5:$E$5,$J$5:$N$5"),
        ([(5, 5), (7, 8), (10, 11)], 5, "$E$5,$G$5:$H$5,$J$5:$K$5"),
    ],
)
def test_range_collection_from_index_column(sheet_module: Sheet, column, n, address):
    rc = RangeCollection.from_index(sheet_module, 5, column)
    assert len(rc) == n
    assert rc.get_address() == address
    assert rc.api.Address == address


def test_range_collection_iter(sheet_module: Sheet):
    rc = RangeCollection.from_index(sheet_module, [(2, 5), (10, 12)], 1)
    for rng, row in zip(rc, [2, 3, 4, 5, 10, 11, 12], strict=True):
        assert rng.row == row


def test_range_collection_first(sheet_module: Sheet):
    rc = RangeCollection.from_index(sheet_module, [(2, 5)], 1)
    cell = rc.first()
    assert cell.row == 2
    assert cell.column == 1


def test_range_collection_repr(sheet_module: Sheet):
    rc = RangeCollection.from_index(sheet_module, [(2, 5), (8, 10)], 5)
    assert repr(rc) == "<RangeCollection $E$2:$E$5,$E$8:$E$10>"
