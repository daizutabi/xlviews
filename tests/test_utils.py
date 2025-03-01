import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.testing import is_app_available


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("ChartType.xlXYScatter", -4169),
        ("BordersIndex.EdgeTop", 8),
        ("Bottom", -4107),
        ("Center", -4108),
        ("Left", -4131),
        ("None", -4142),
        ("Right", -4152),
        ("Top", -4160),
    ],
)
def test_constant(name: str, value: int):
    from xlviews.utils import constant

    assert constant(name) == value
    assert constant(*name.split(".")) == value


@pytest.mark.parametrize(
    ("columns", "lst"),
    [
        ("B", ["B"]),
        (["A", "C"], ["A", "C"]),
        (":B", ["A", "B"]),
        (["::B", "C"], ["A", "C"]),
    ],
)
def test_iter_columns(columns, lst):
    from xlviews.utils import iter_columns

    df = DataFrame([[1, 2, 3], [4, 5, 6], [7, 8, 9]], columns=["A", "B", "C"])
    assert list(iter_columns(df, columns)) == lst


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_validate_list(sheet: Sheet):
    from xlviews.utils import add_validate_list

    rng = sheet.range("a1")
    add_validate_list(rng, [1, 2, 3], 2)
    assert rng.value == 2

    assert rng.api.Validation.Type == 3
    assert rng.api.Validation.Operator == 3
    assert rng.api.Validation.Formula1 == "1,2,3"
