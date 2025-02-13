import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def test_add_column(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    rng = sf.add_column("c", autofit=True, style=True)
    assert rng.get_address() == "$E$3:$E$5"
    assert sf.headers == [None, "a", "b", "c"]


def test_add_column_value(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    rng = sf.add_column("c", [7, 8, 9], number_format="0.0")
    assert rng.get_address() == "$E$3:$E$5"
    np.testing.assert_array_equal(sf.data["c"], [7, 8, 9])


@pytest.mark.parametrize(
    ("formula", "value"),
    [("={a}+{b}", [6, 8, 10, 12]), ("={a}*{b}", [5, 12, 21, 32])],
)
@pytest.mark.parametrize("use_setitem", [False, True])
def test_add_formula_column(formula, value, use_setitem, sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(2, 3, data=df, sheet=sheet)
    if use_setitem:
        sf["c"] = formula
    else:
        sf.add_formula_column("c", formula)

    np.testing.assert_array_equal(sf.data["c"], value)

    sf.add_formula_column("c", formula + "+1")
    np.testing.assert_array_equal(sf.data["c"], [v + 1 for v in value])


@pytest.mark.parametrize("apply", [lambda x: x, Range.from_range])
def test_add_formula_column_range(apply, sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(2, 3, data=df, sheet=sheet)
    rng = apply(sf.add_column("c"))
    sf.add_formula_column(rng, "={b}-{a}", number_format="0", autofit=True, style=True)
    np.testing.assert_array_equal(rng.value, [4, 4, 4, 4])


@pytest.mark.parametrize(
    ("formula", "value"),
    [
        ("={a}+{b}+{c}", ([7, 9, 11, 13], [10, 12, 14, 16])),
        ("={a}*{b}*{c}", ([5, 12, 21, 32], [20, 48, 84, 128])),
    ],
)
def test_formula_wide(formula, value, sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(10, 3, data=df, sheet=sheet)
    sf.add_wide_column("c", [1, 2, 3, 4], number_format="0", style=True)
    sf.add_formula_column("c", formula, number_format="0", autofit=True)
    a = sf.range(("c", 1)).value
    np.testing.assert_array_equal(a, value[0])
    a = sf.range(("c", 4)).value
    np.testing.assert_array_equal(a, value[1])
