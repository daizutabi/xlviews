import pytest
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.range.range import Range
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module", params=["A1", "A1:A3", "F4:I4", "C1:E3"])
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


def test_range_str_str(addr_impl: str, include_sheetname, external):
    rng = Range(addr_impl, addr_impl)
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


@pytest.fixture(scope="module")
def rng(rng_impl: RangeImpl):
    return Range(rng_impl)


def test_repr(rng: Range, rng_impl: RangeImpl):
    assert repr(rng) == repr(rng_impl)


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
