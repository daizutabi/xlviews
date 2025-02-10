import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_excel_installed
from xlviews.testing.heat_frame.facet import Facet

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Facet(sheet_module)


@pytest.fixture(scope="module")
def df(fc: Facet):
    return fc.df


# def test_group(df: DataFrame):
#     df = df.pivot_table("v", ["Y", "y"], ["X", "x"], aggfunc=lambda x: x)
#     print(df)
# assert 0
