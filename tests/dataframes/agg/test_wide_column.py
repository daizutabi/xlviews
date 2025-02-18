import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import WideColumn

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


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
    return sf.value


@pytest.mark.parametrize("func", ["sum", "max", "mean"])
def test_str(sf: SheetFrame, df: DataFrame, func: str):
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    assert isinstance(a, Series)
    sf = SheetFrame(20, 2, a.to_frame(), sf.sheet)
    np.testing.assert_array_equal(sf.value[func], b)


def test_list(sf: SheetFrame, df: DataFrame):
    func = ["min", "max", "median", "sum"]
    a = sf.agg(func, formula=True)
    b = df.agg(func)  # type: ignore
    assert isinstance(a, DataFrame)
    assert isinstance(b, DataFrame)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(50, 2, a, sf.sheet)
    np.testing.assert_array_equal(sf.value, b)


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
