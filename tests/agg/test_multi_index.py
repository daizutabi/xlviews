import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.frame import SheetFrame
from xlviews.utils import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 2, 2, 2, 2],
            "y": [1, 1, 2, 2, 1, 1, 2, 2],
            "a": [1, 2, 3, 4, 5, 6, 7, 8],
            "b": [11, 12, 13, 14, 15, 16, 17, 18],
        },
    )
    return df.set_index(["x", "y"])


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(2, 2, data=df, style=False, sheet=sheet_module)


@pytest.mark.parametrize(
    ("func", "value"),
    [("sum", 36), ("min", 1), ("max", 8), ("mean", 4.5)],
)
def test_df_str(df: DataFrame, func: str, value: float):
    s = df.agg(func)
    assert isinstance(s, Series)
    assert s.index.to_list() == ["a", "b"]
    assert s["a"] == value


def test_df_dict(df: DataFrame):
    s = df.agg({"a": "min", "b": "max"})
    assert isinstance(s, Series)
    assert s.index.to_list() == ["a", "b"]
    np.testing.assert_array_equal(s, [1, 18])


def test_df_dict_one(df: DataFrame):
    s = df.agg({"a": "min"})
    assert isinstance(s, Series)
    assert s.index.to_list() == ["a"]
    np.testing.assert_array_equal(s, [1])


def test_df_list(df: DataFrame):
    x = df.agg(["min", "max"])
    assert isinstance(x, DataFrame)
    assert x.index.to_list() == ["min", "max"]
    assert x.columns.to_list() == ["a", "b"]
    np.testing.assert_array_equal(x, [[1, 11], [8, 18]])


@pytest.mark.parametrize("func", ["sum", "count", "min", "max", "mean"])
def test_sf_str(sf: SheetFrame, df: DataFrame, func: str):
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data[0], b)


def test_sf_str_columns(sf: SheetFrame):
    a = sf.agg("mean", columns="a", formula=True)
    assert len(a) == 1
    sf = SheetFrame(20, 10, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data, [[4.5]])


def test_sf_dict(sf: SheetFrame, df: DataFrame):
    func = {"a": "min", "b": "max"}
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data[0], b)


def test_sf_list(sf: SheetFrame, df: DataFrame):
    func = ["min", "max"]
    a = sf.agg(func, formula=True)
    b = df.agg(func)  # type: ignore
    assert isinstance(a, DataFrame)
    assert a.index.to_list() == b.index.to_list()
    assert a.columns.to_list() == b.columns.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data, b)


def test_sf_list_columns(sf: SheetFrame, df: DataFrame):
    a = sf.agg(["sum", "count"], columns="b", formula=True)
    assert isinstance(a, DataFrame)
    sf = SheetFrame(20, 20, data=a, sheet=sf.sheet, style=False)
    np.testing.assert_array_equal(sf.data, [[116], [8]])


def test_sf_none(sf: SheetFrame):
    s = sf.agg(None)
    assert isinstance(s, Series)
    assert s["a"] == "$D$3:$D$10"
    assert s["b"] == "$E$3:$E$10"


def test_sf_first(sf: SheetFrame):
    s = sf.agg("first", formula=True)
    assert isinstance(s, Series)
    assert s["a"] == "=$D$3"
    assert s["b"] == "=$E$3"


def test_sf_first_none(sf: SheetFrame):
    s = sf.agg({"x": "first", "b": None}, formula=True)
    assert isinstance(s, Series)
    assert len(s) == 2
    assert s["x"] == "=$B$3"
    assert s["b"] == "=$E$3:$E$10"


def test_group(df: DataFrame):
    print(df.groupby("x").agg("sum"))
