import pytest
from xlwings import Sheet


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

    with pytest.raises(ValueError, match="sheet is required when cell is a tuple"):
        reference((4, 5))
