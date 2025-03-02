import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.core.formula import NONCONST_VALUE
from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 1, 1, 1], "b": [2, 2, 3, 3], "c": [4, 4, 4, 4]})


@pytest.fixture(scope="module", params=[RangeImpl, Range], ids=["RangeImpl", "Range"])
def rng(df: DataFrame, sheet_module: Sheet, request: pytest.FixtureRequest):
    rng = sheet_module.range("B3")
    rng.options(DataFrame, header=1, index=False).value = df
    rng: RangeImpl = rng.expand()
    return request.param(
        (rng.row, rng.column),
        (rng.last_cell.row, rng.last_cell.column),
    )


def test_range_address(rng: Range | RangeImpl):
    assert rng.get_address() == "$B$3:$D$7"


@pytest.fixture(scope="module")
def column(rng: Range | RangeImpl):
    start = rng[0].offset(1)
    end = start.offset(3)
    return rng.__class__((start.row, start.column), (end.row, end.column))


def test_column(column: Range | RangeImpl):
    assert column.get_address() == "$B$4:$B$7"


@pytest.fixture(scope="module")
def const_header(rng: Range | RangeImpl):
    start = (rng[0].row, rng[0].column)
    rng_impl = rng if isinstance(rng, RangeImpl) else rng.impl
    end = rng_impl[0].expand("right")
    return rng.__class__(start, (end[-1].row, end[-1].column)).offset(-1)


def test_header(const_header: Range | RangeImpl):
    assert const_header.get_address() == "$B$2:$D$2"


@pytest.mark.parametrize(("k", "value"), [(0, 1), (1, NONCONST_VALUE), (2, 4)])
def test_const(column: Range | RangeImpl, const_header: Range | RangeImpl, k, value):
    from xlviews.core.formula import const

    const_header.value = "=" + const(column)
    assert const_header.value[k] == value


@pytest.fixture(scope="module")
def ranges(sheet_module: Sheet):
    rng = sheet_module.range("B100")
    rng.options(transpose=True).value = [1, 2, 3, 4, 0, 5, 6, 7, 8, 9, 10]
    rng1 = rng.expand("down")
    sheet_module.range("B104").value = None

    rng = sheet_module.range("C100")
    rng.options(transpose=True).value = [11, 12, 13, 14, 15, 16, 17, 0, 18, 19, 20]
    rng2 = rng.expand("down")
    sheet_module.range("C107").value = None

    return [Range.from_range(rng1), Range.from_range(rng2)]


def test_ranges(ranges: list[Range]):
    assert ranges[0].get_address() == "$B$100:$B$110"
    assert ranges[1].get_address() == "$C$100:$C$110"


def test_aggregate_value(ranges: list[Range]):
    from xlviews.core.formula import aggregate

    x = aggregate("count", ranges)
    assert x == "AGGREGATE(2,7,$B$100:$B$110,$C$100:$C$110)"


def test_aggregate_none(ranges: list[Range]):
    from xlviews.core.formula import aggregate

    x = aggregate(None, ranges)
    assert x == "$B$100:$B$110,$C$100:$C$110"


def test_aggregate_formula(ranges: list[Range]):
    from xlviews.core.formula import aggregate

    x = aggregate("max", ranges, formula=True)
    assert x == "=AGGREGATE(4,7,$B$100:$B$110,$C$100:$C$110)"


def test_aggreate_range_range(ranges: list[Range]):
    from xlviews.core.formula import aggregate

    assert aggregate("max", ranges[0]) == "AGGREGATE(4,7,$B$100:$B$110)"


FUNC_VALUES = [
    ("count", 20),
    ("sum", 210),
    ("min", 1),
    ("max", 20),
    ("mean", 10.5),
    ("median", 10.5),
    ("std", np.std(range(1, 21))),
    ("soa", np.std(range(1, 21)) / 10.5),
]


@pytest.mark.parametrize(("func", "value"), FUNC_VALUES)
def test_aggregate_str(ranges: list[Range], func, value):
    from xlviews.core.formula import aggregate

    cell = ranges[0].sheet.range("D100")
    cell.value = aggregate(func, ranges, formula=True)
    assert cell.value == value


@pytest.mark.parametrize(("func", "value"), FUNC_VALUES)
@pytest.mark.parametrize("apply", [lambda x: x, Range.from_range])
def test_aggregate_range(ranges: list[Range], func, value, apply):
    from xlviews.core.formula import aggregate

    ref = apply(ranges[0].sheet.range("E100"))
    ref.value = func
    formula = aggregate(ref, ranges)
    cell = ranges[0].sheet.range("D100")
    cell.value = "=" + formula
    assert cell.value == value


def test_aggreate_func_invalid():
    from xlviews.core.formula import aggregate

    with pytest.raises(ValueError, match="Invalid aggregate function: sin"):
        aggregate("sin", "A1:A10")
