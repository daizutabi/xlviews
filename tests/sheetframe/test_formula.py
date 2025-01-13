import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def sf1(sheet_module: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(2, 3, data=df, style=False, sheet=sheet_module)


@pytest.mark.parametrize(
    ("formula", "value"),
    [
        ("={a}+{b}", [6, 8, 10, 12]),
        ("={a}*{b}", [5, 12, 21, 32]),
        ("={a}-{b}", [-4, -4, -4, -4]),
    ],
)
@pytest.mark.parametrize("use_setitem", [False, True])
def test_formula(sf1: SheetFrame, formula, value, use_setitem):
    if use_setitem:
        sf1["c"] = formula
    else:
        sf1.add_formula_column("c", formula)

    np.testing.assert_array_equal(sf1["c"], value)


@pytest.fixture(scope="module")
def sf2(sheet_module: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(10, 3, data=df, style=False, sheet=sheet_module)
    sf.add_wide_column("c", [1, 2, 3, 4])
    return sf


@pytest.mark.parametrize(
    ("formula", "value"),
    [
        ("={a}+{b}+{c}", ([7, 9, 11, 13], [10, 12, 14, 16])),
        ("={a}*{b}*{c}", ([5, 12, 21, 32], [20, 48, 84, 128])),
    ],
)
def test_formula_wide(sf2: SheetFrame, formula, value):
    sf2.add_formula_column("c", formula, number_format="0", autofit=True)
    np.testing.assert_array_equal(sf2[("c", 1)], value[0])
    np.testing.assert_array_equal(sf2[("c", 4)], value[1])
