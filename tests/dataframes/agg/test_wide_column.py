import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame import WideColumn

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return WideColumn(sheet_module, 4, 2)


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    sf = fc.sf
    sf.add_formula_column("u", "={u}+{a}")
    sf.add_formula_column("v", "={v}+{b}")
    return sf


@pytest.fixture(scope="module")
def df(sf: SheetFrame):
    return sf.data


def test_sf_data(df: DataFrame):
    print(df)
    assert 0


@pytest.mark.parametrize("func", ["sum", "max", "mean"])
def test_str(sf: SheetFrame, df: DataFrame, func: str):
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    print(a)
    print(b)
    assert 0
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data[0], b)
