import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": ["a"] * 8 + ["b"] * 8 + ["a"] * 4,
            "y": (["c"] * 4 + ["d"] * 4) * 2 + ["c"] * 4,
            "z": range(1, 21),
            "a": range(20),
            "b": list(range(10)) + list(range(0, 30, 3)),
            "c": list(range(20, 40, 2)) + list(range(0, 20, 2)),
        },
    )
    df = df.set_index(["x", "y", "z"])
    df.iloc[[4, -1], 0] = np.nan
    df.iloc[[3, 6, 9], -1] = np.nan
    return df


@pytest.fixture(scope="module")
def sf_parent(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(3, 3, data=df, table=True, sheet=sheet_module)
