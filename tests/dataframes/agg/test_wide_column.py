import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

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


@pytest.mark.parametrize("func", ["sum", "max", "mean"])
def test_str(sf: SheetFrame, df: DataFrame, func: str):
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data[0], b)


def test_list(sf: SheetFrame, df: DataFrame):
    func = ["min", "max", "median", "sum"]
    a = sf.agg(func, formula=True)
    b = df.agg(func)  # type: ignore
    assert isinstance(a, DataFrame)
    assert a.index.to_list() == b.index.to_list()
    assert a.columns.to_list() == b.columns.to_list()
    sf = SheetFrame(50, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data, b)


def test_sf_none(sf: SheetFrame):
    s = sf.agg(None)
    assert isinstance(s, Series)
    assert s["a"] == "$D$5:$D$9"
    assert s["b"] == "$E$5:$E$9"


def test_sf_first(sf: SheetFrame):
    s = sf.agg("first", formula=True)
    assert isinstance(s, Series)
    assert s["a"] == "=$D$5"
    assert s["b"] == "=$E$5"


@pytest.mark.parametrize("func", ["sum", "median", "mean"])
@pytest.mark.parametrize("by", ["x", "y"])
def test_sf_group_str_str(sf: SheetFrame, df: DataFrame, func, by):
    a = sf.groupby(by).agg(func, as_address=True, formula=True)
    b = df.groupby(by).agg(func).astype(float)
    sf = SheetFrame(50, 30, data=a, sheet=sf.sheet, style=False)
    assert sf.data.equals(b)
