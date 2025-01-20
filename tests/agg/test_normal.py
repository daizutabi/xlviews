import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.frame import SheetFrame
from xlviews.table import Table
from xlviews.utils import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def df():
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]}, index=["x", "x", "y", "y"])
    df.index.name = "name"
    return df


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(2, 3, data=df, style=False, sheet=sheet_module)
