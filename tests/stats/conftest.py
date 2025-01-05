import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": ["a"] * 10 + ["b"] * 10,
            "y": (["c"] * 6 + ["d"] * 4) * 2,
            "z": range(1, 21),
            "a": range(20),
            "b": list(range(10)) + list(range(0, 30, 3)),
            "c": list(range(20, 40, 2)) + list(range(0, 20, 2)),
        },
    )
    return df.set_index(["x", "y", "z"])


@pytest.fixture(scope="module")
def sf_parent(df: DataFrame, sheet_module: Sheet):
    df = df.copy()
    df.iloc[[4, -1], 0] = np.nan
    df.iloc[[3, 6, 9], -1] = np.nan
    return SheetFrame(sheet_module, 2, 3, data=df, table=True)
