import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


def test_add_column(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    sf.add_column("c", autofit=True, style=True)
    assert sf.columns.to_list() == ["a", "b", "c"]


def test_add_column_value(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet)
    sf.add_column("c", [7, 8, 9], number_format="0.0")
    assert sheet.range("E3:E5").value == [7, 8, 9]


@pytest.mark.parametrize(
    ("formula", "value"),
    [("={a}+{b}", [6, 8, 10, 12]), ("={a}*{b}", [5, 12, 21, 32])],
)
def test_add_formula_column(formula, value, sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(2, 3, data=df, sheet=sheet)
    sf.add_formula_column("c", formula)
    assert sheet.range("F3:F6").value == value

    sf.add_formula_column("c", formula + "+1", style=True)
    assert sheet.range("F3:F6").value == [v + 1 for v in value]


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
    assert sheet.range((11, 6), (14, 6)).value == value[0]
    assert sheet.range((11, 9), (14, 9)).value == value[1]
