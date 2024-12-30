import pytest
from xlwings import Sheet


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
        ([3, 8], 6, "$E$3:$E$8"),  # TODO: Delete
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
        ([3, 8], 6, "$C$10:$H$10"),  # TODO: Delete
    ],
)
def test_multirange_column(sheet_module: Sheet, index, n, rng):
    from xlviews.range import multirange

    x = multirange(sheet_module, 10, index)
    assert len(x) == n
    assert x.get_address() == rng


def test_reference_str(sheet_module: Sheet):
    from xlviews.range import reference

    assert reference(sheet_module, "x") == "x"


def test_reference_range(sheet_module: Sheet):
    from xlviews.range import reference

    cell = sheet_module.range(4, 5)

    ref = reference(sheet_module, cell)
    assert ref == f"={sheet_module.name}!$E$4"
