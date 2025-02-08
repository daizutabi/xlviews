import pytest
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.range.range import Range
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module", params=["C5", "D10:D13", "F4:I4", "C5:E9"])
def addr(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def rng_impl(addr, sheet_module: Sheet):
    return sheet_module.range(addr)


@pytest.fixture(scope="module", params=[True, False])
def include_sheetname(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module", params=[True, False])
def external(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def addr_impl(rng_impl: RangeImpl, include_sheetname, external):
    return rng_impl.get_address(include_sheetname=include_sheetname, external=external)


def test_range_str(addr_impl: str, include_sheetname, external):
    rng = Range(addr_impl)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


@pytest.mark.parametrize("func", [lambda x: x, Range])
def test_range_str_str(addr_impl: str, include_sheetname, external, func):
    rng = func(Range(addr_impl, addr_impl))
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_range(rng_impl: RangeImpl, addr_impl, include_sheetname, external):
    rng = Range(rng_impl)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_range_range(rng_impl: RangeImpl, addr_impl, include_sheetname, external):
    rng = Range(rng_impl, rng_impl.last_cell)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_tuple(rng_impl: RangeImpl, include_sheetname, external):
    rng = Range((rng_impl.row, rng_impl.column))
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    y = rng_impl[0].get_address(include_sheetname=include_sheetname, external=external)
    assert x == y


def test_range_tuple_tuple(rng_impl: RangeImpl, addr_impl, include_sheetname, external):
    cell1 = (rng_impl.row, rng_impl.column)
    cell2 = (rng_impl.last_cell.row, rng_impl.last_cell.column)
    rng = Range(cell1, cell2)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_error(sheet: Sheet, sheet_module: Sheet):
    with pytest.raises(ValueError, match="Cells are not in the same sheet"):
        Range(sheet.range("A1"), sheet_module.range("B2"))


def test_range_book_error(sheet_module: Sheet):
    with pytest.raises(ValueError, match="Book name does not match"):
        Range("[a]b!A1")


def test_range_sheet_error(sheet_module: Sheet):
    with pytest.raises(ValueError, match="Sheet not found"):
        Range("b!A1")


def test_range_type_error(sheet_module: Sheet):
    with pytest.raises(TypeError, match="Invalid type"):
        Range(1)  # type: ignore


def test_range_other_sheet(sheet_module: Sheet, sheet: Sheet):
    rng = Range(f"{sheet.name}!A1", sheet=sheet_module)
    assert rng.sheet.name == sheet.name


@pytest.fixture(scope="module")
def rng(rng_impl: RangeImpl):
    return Range(rng_impl)


def test_len(rng: Range, rng_impl: RangeImpl):
    assert len(rng) == len(rng_impl)


def test_iter(rng: Range, rng_impl: RangeImpl):
    x = [r.get_address() for r in rng]
    y = [r.get_address() for r in rng_impl]
    assert x == y


def test_getitem(rng: Range, rng_impl: RangeImpl):
    for k in range(len(rng)):
        assert rng[k].get_address() == rng_impl[k].get_address()


def test_getitem_neg(rng: Range, rng_impl: RangeImpl):
    for k in range(len(rng)):
        assert rng[-k].get_address() == rng_impl[-k].get_address()


def test_getitem_error(rng: Range):
    with pytest.raises(IndexError, match="Index out of range"):
        rng[100]


def test_repr(rng: Range, rng_impl: RangeImpl):
    assert repr(rng) == repr(rng_impl)


@pytest.mark.parametrize("row_offset", [2, 0, -2])
@pytest.mark.parametrize("column_offset", [2, 0, -2])
def test_offset(rng: Range, rng_impl: RangeImpl, row_offset, column_offset):
    x = rng.offset(row_offset, column_offset)
    y = rng_impl.offset(row_offset, column_offset)
    assert x.get_address() == y.get_address()


def test_impl_from(rng: Range, rng_impl: RangeImpl):
    rng_impl.value = rng_impl.get_address(external=True)
    assert rng_impl.value == rng.impl.value


def test_impl_to(rng: Range, rng_impl: RangeImpl):
    rng.impl.value = rng.get_address()
    assert rng_impl.value == rng.impl.value


def test_iter_addresses(rng: Range, rng_impl: RangeImpl, external):
    x = list(rng.iter_addresses(external=external))
    y = [r.get_address(external=external) for r in rng_impl]
    assert x == y


def test_iter_addresses_formula(rng: Range, rng_impl: RangeImpl, external):
    x = list(rng.iter_addresses(external=external, formula=True))
    y = ["=" + r.get_address(external=external) for r in rng_impl]
    assert x == y


@pytest.mark.parametrize(
    ("addr", "value"),
    [
        ("A1", "a"),
        ("A1:C1", (("a", "a", "a"),)),
        ("A1:A3", (("a",), ("a",), ("a",))),
        ("A1:B2", (("a", "a"), ("a", "a"))),
    ],
)
def test_api(addr, value, sheet: Sheet):
    rng = Range(addr, sheet=sheet)
    rng.impl.value = "a"
    assert rng.api.Value == value
