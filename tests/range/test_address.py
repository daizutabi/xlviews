import pytest
from xlwings import Sheet

from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def test_reference_str(sheet_module: Sheet):
    from xlviews.range.address import reference

    assert reference("x", sheet_module) == "x"


def test_reference_range(sheet_module: Sheet):
    from xlviews.range.address import reference

    cell = sheet_module.range(4, 5)

    ref = reference(cell)
    assert ref == f"={sheet_module.name}!$E$4"


def test_reference_tuple(sheet_module: Sheet):
    from xlviews.range.address import reference

    ref = reference((4, 5), sheet_module)
    assert ref == f"={sheet_module.name}!$E$4"


def test_reference_error(sheet_module: Sheet):
    from xlviews.range.address import reference

    m = "`sheet` is required when `cell` is a tuple"
    with pytest.raises(ValueError, match=m):
        reference((4, 5))


@pytest.mark.parametrize(
    ("rng", "addrs"),
    [("A1", ["A1"]), ("A1:A3", ["A1:A3"])],
)
def test_iter_addresses(rng, addrs, sheet_module: Sheet):
    from xlviews.range.address import iter_addresses

    if isinstance(rng, str):
        rngs = sheet_module.range(rng)
    else:
        rngs = [sheet_module.range(r) for r in rng]

    x = list(iter_addresses(rngs, row_absolute=False, column_absolute=False))
    assert x == addrs
