import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.range.formula import NONCONST_VALUE
from xlviews.range.range_collection import RangeCollection
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 1, 1, 1], "b": [2, 2, 3, 3], "c": [4, 4, 4, 4]})


@pytest.fixture(scope="module")
def rng_impl(df: DataFrame, sheet_module: Sheet):
    rng = sheet_module.range("B3")
    rng.options(DataFrame, header=1, index=False).value = df
    return rng.expand()


def test_range_address(rng_impl: RangeImpl):
    assert rng_impl.get_address() == "$B$3:$D$7"


@pytest.mark.parametrize("k", range(3))
def test_range_value(rng_impl: RangeImpl, df: DataFrame, k: int):
    value = rng_impl.options(transpose=True).value
    assert isinstance(value, list)
    assert len(value) == 3
    np.testing.assert_array_equal(df[value[k][0]], value[k][1:])


@pytest.fixture(scope="module")
def column(rng: Range):
    start = rng[0].offset(1)
    end = start.offset(3)
    return rng.sheet.range(start, end)


def test_column(column: Range):
    assert column.get_address() == "$B$4:$B$7"


# @pytest.fixture(scope="module")
# def const_header(rng: Range):
#     end = rng[0].expand("right")
#     return rng.sheet.range(rng[0], end).offset(-1)


# def test_header(const_header: Range):
#     assert const_header.get_address() == "$B$2:$D$2"


# @pytest.mark.parametrize(("k", "value"), [(0, 1), (1, NONCONST_VALUE), (2, 4)])
# def test_const(column: Range, const_header: Range, k, value):
#     from xlviews.range.formula import const

#     const_header.value = const(column, "=")
#     assert const_header.value[k] == value


# @pytest.fixture(scope="module")
# def ranges(sheet_module: Sheet):
#     rng = sheet_module.range("B100")
#     rng.options(transpose=True).value = [1, 2, 3, 4, 0, 5, 6, 7, 8, 9, 10]
#     rng1 = rng.expand("down")
#     sheet_module.range("B104").value = None

#     rng = sheet_module.range("C100")
#     rng.options(transpose=True).value = [11, 12, 13, 14, 15, 16, 17, 0, 18, 19, 20]
#     rng2 = rng.expand("down")
#     sheet_module.range("C107").value = None

#     return [rng1, rng2]


# def test_ranges(ranges: list[Range]):
#     assert ranges[0].get_address() == "$B$100:$B$110"
#     assert ranges[1].get_address() == "$C$100:$C$110"


# @pytest.mark.parametrize("apply", [list, RangeCollection])
# def test_aggregate_value(ranges: list[Range], apply):
#     from xlviews.range.formula import aggregate

#     x = aggregate("count", apply(ranges))
#     assert x == "AGGREGATE(2,7,$B$100:$B$110,$C$100:$C$110)"


# @pytest.mark.parametrize("apply", [list, RangeCollection])
# def test_aggregate_none(ranges: list[Range], apply):
#     from xlviews.range.formula import aggregate

#     x = aggregate(None, apply(ranges))
#     assert x == "$B$100:$B$110,$C$100:$C$110"


# @pytest.mark.parametrize("apply", [list, RangeCollection])
# def test_aggregate_formula(ranges: list[Range], apply):
#     from xlviews.range.formula import aggregate

#     x = aggregate("max", apply(ranges), formula=True)
#     assert x == "=AGGREGATE(4,7,$B$100:$B$110,$C$100:$C$110)"


# FUNC_VALUES = [
#     ("count", 20),
#     ("sum", 210),
#     ("min", 1),
#     ("max", 20),
#     ("mean", 10.5),
#     ("median", 10.5),
#     ("std", np.std(range(1, 21))),
#     ("soa", np.std(range(1, 21)) / 10.5),
# ]


# @pytest.mark.parametrize(("func", "value"), FUNC_VALUES)
# @pytest.mark.parametrize("apply", [list, RangeCollection])
# def test_aggregate_str(ranges: list[Range], apply, func, value):
#     from xlviews.range.formula import aggregate

#     formula = aggregate(func, apply(ranges))
#     cell = ranges[0].sheet.range("D100")
#     cell.value = "=" + formula
#     assert cell.value == value


# @pytest.mark.parametrize(("func", "value"), FUNC_VALUES)
# @pytest.mark.parametrize("apply", [list, RangeCollection])
# def test_aggregate_range(ranges: list[Range], apply, func, value):
#     from xlviews.range.formula import aggregate

#     ref = ranges[0].sheet.range("E100")
#     ref.value = func
#     formula = aggregate(ref, apply(ranges))
#     cell = ranges[0].sheet.range("D100")
#     cell.value = "=" + formula
#     assert cell.value == value
