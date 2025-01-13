import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    return df.set_index(["x", "y"])


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(3, 2, data=df, style=False, sheet=sheet_module)
