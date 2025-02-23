import pytest
from pandas import DataFrame

# from xlwings import Sheet

# from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def df():
    return DataFrame(
        {"x": [1, 1, 1, 2, 2, 2, 3, 3, 3], "y": [4, 5, 6, 4, 5, 6, 4, 5, 6]},
    ).set_index("x")


@pytest.mark.parametrize(
    ("style", "by"),
    [
        (None, []),
        ("x", ["x"]),
        ("y", ["y"]),
        ("z", []),
        (["x", "y"], ["x", "y"]),
        (["x", "y", "z"], ["x", "y"]),
        ({"x": "x", "z": "z"}, ["x"]),
        ({("y", "z"): "y", "x": "x"}, ["y", "x"]),
        ({"z": "z"}, []),
    ],
)
def test_iter_by(df: DataFrame, style, by):
    from xlviews.chart.plot import iter_by

    assert list(iter_by(df, style)) == by


@pytest.mark.parametrize(
    ("styles", "by"),
    [
        (["x"], ["x"]),
        (["y", ["x", "y"]], ["x", "y"]),
        ([{("y", "z"): "y"}, {"x": "x"}], ["x", "y"]),
    ],
)
def test_get_by(df: DataFrame, styles, by):
    from xlviews.chart.plot import get_by

    assert get_by(df, styles) == by
