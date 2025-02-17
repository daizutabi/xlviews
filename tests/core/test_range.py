import numpy as np
import pytest
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


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


def test_range_int_int(rng_impl: RangeImpl, include_sheetname, external):
    rng = Range(rng_impl.row, rng_impl.column)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    y = rng_impl[0].get_address(include_sheetname=include_sheetname, external=external)
    assert x == y


def test_range_tuple_first(rng_impl: RangeImpl, include_sheetname, external):
    rng = Range((rng_impl.row, rng_impl.column))
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    y = rng_impl[0].get_address(include_sheetname=include_sheetname, external=external)
    assert x == y


def test_range_tuple_last(rng_impl: RangeImpl, include_sheetname, external):
    rng = Range((rng_impl.last_cell.row, rng_impl.last_cell.column))
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    y = rng_impl[-1].get_address(include_sheetname=include_sheetname, external=external)
    assert x == y


def test_range_tuple_tuple(rng_impl: RangeImpl, addr_impl, include_sheetname, external):
    cell1 = (rng_impl.row, rng_impl.column)
    cell2 = (rng_impl.last_cell.row, rng_impl.last_cell.column)
    rng = Range(cell1, cell2)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_from_range(rng_impl: RangeImpl, addr_impl, include_sheetname, external):
    rng = Range.from_range(rng_impl)
    x = rng.get_address(include_sheetname=include_sheetname, external=external)
    assert x == addr_impl


def test_range_from_range_first(rng_impl: RangeImpl):
    rng = Range.from_range(rng_impl[0])
    assert rng.get_address() == rng_impl[0].get_address()


def test_range_from_range_last(rng_impl: RangeImpl):
    rng = Range.from_range(rng_impl.last_cell)
    assert rng.get_address() == rng_impl.last_cell.get_address()


def test_range_error_tuple(sheet_module: Sheet):
    with pytest.raises(TypeError, match="cell2 must be a tuple or None"):
        Range((1, 1), 1)


def test_range_error_int(sheet_module: Sheet):
    with pytest.raises(TypeError, match="cell2 must be an integer"):
        Range(1, (1, 1))


@pytest.fixture(scope="module")
def rng(rng_impl: RangeImpl):
    return Range.from_range(rng_impl)


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


def test_last_cell(rng: Range, rng_impl: RangeImpl):
    assert rng.last_cell.get_address() == rng_impl.last_cell.get_address()


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


def test_frame_range_cell(sheet_module: Sheet):
    rng = Range((1, 1), (1, 1), sheet_module).frame
    df = rng.get_address()
    assert df.shape == (1, 1)
    assert df.index.tolist() == [0]
    assert df.columns.tolist() == [0]
    np.testing.assert_array_equal(df, [["$A$1"]])


def test_frame_range_row(sheet_module: Sheet):
    rng = Range((1, 1), (1, 3), sheet_module).frame
    df = rng.get_address(row_absolute=False)
    assert df.shape == (1, 3)
    assert df.index.tolist() == [0]
    assert df.columns.tolist() == [0, 1, 2]
    np.testing.assert_array_equal(df, [["$A1", "$B1", "$C1"]])


def test_frame_range_column(sheet_module: Sheet):
    rng = Range((1, 1), (3, 1), sheet_module).frame
    df = rng.get_address(column_absolute=False)
    assert df.shape == (3, 1)
    assert df.index.tolist() == [0, 1, 2]
    assert df.columns.tolist() == [0]
    np.testing.assert_array_equal(df, [["A$1"], ["A$2"], ["A$3"]])


def test_frame_range_matrix(sheet_module: Sheet):
    rng = Range((1, 1), (3, 2), sheet_module).frame
    df = rng.get_address(column_absolute=False, row_absolute=False, formula=True)
    assert df.shape == (3, 2)
    assert df.index.tolist() == [0, 1, 2]
    assert df.columns.tolist() == [0, 1]
    np.testing.assert_array_equal(df, [["=A1", "=B1"], ["=A2", "=B2"], ["=A3", "=B3"]])


def test_frame_range_cell_sheetname(sheet_module: Sheet):
    rng = Range((1, 1), (1, 1), sheet_module).frame
    df = rng.get_address(include_sheetname=True)
    assert df.loc[0, 0] == f"{sheet_module.name}!$A$1"


def test_frame_range_cell_external(sheet_module: Sheet):
    rng = Range((1, 1), (1, 1), sheet_module).frame
    df = rng.get_address(external=True)
    assert df.loc[0, 0] == f"[{sheet_module.book.name}]{sheet_module.name}!$A$1"
