import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 2, 3, data=df, style=False)


# def test_formula(sf: SheetFrame):
#     rng=self.
#     sf.add_formula_column("a", "=A1+B1", "c")
#     assert sf.cell.value == [
#         [None, "a", "b", "c"],
#         [0, 1, 5, 6],
#         [1, 2, 6, 8],
#         [2, 3, 7, 10],
#         [3, 4, 8, 12],
#     ]
