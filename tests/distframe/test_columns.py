import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.distframe import DistFrame
from xlviews.formula import NONCONST_VALUE
from xlviews.sheetframe import SheetFrame


@pytest.mark.parametrize(
    ("columns", "values"),
    [
        (None, ["x", "y", "a_n", "a_v", "a_s", "b_n", "b_v", "b_s"]),
        ("a", ["x", "y", "a_n", "a_v", "a_s"]),
        ("b", ["x", "y", "b_n", "b_v", "b_s"]),
    ],
)
def test_columns(df: DataFrame, sheet: Sheet, columns, values):
    sf = SheetFrame(3, 2, data=df, style=False, sheet=sheet)
    sf = DistFrame(sf, columns, by=["x", "y"])
    assert sf.columns == values


def test_group_error(sheet: Sheet):
    df = DataFrame(
        {
            "x": [1, 1, 1, 2, 2, 2, 1, 1, 1, 2, 2, 2, 2, 2],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    df = df.set_index(["x", "y"])

    sf = SheetFrame(3, 2, data=df, style=False, sheet=sheet)
    with pytest.raises(ValueError, match="group must be continuous"):
        sf = DistFrame(sf, None, by=["x", "y"])


def test_group_const(sheet: Sheet):
    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    df = df.set_index(["x", "y"])

    sf = SheetFrame(3, 2, data=df, style=False, sheet=sheet)
    sf = sf.distframe(by=["x", "y"])

    assert sf.sheet["H20"].value == 1
    assert sf.sheet["H21"].value == NONCONST_VALUE
