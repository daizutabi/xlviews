import numpy as np
import pytest
from pandas import DataFrame, Index, MultiIndex, Series
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


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


@pytest.fixture
def sf(df: DataFrame, sheet: Sheet):
    return SheetFrame(2, 2, data=df, sheet=sheet)


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
    sf = SheetFrame(20, 2, data=a.to_frame(), sheet=sf.sheet)
    np.testing.assert_array_equal(sf.value[func], b)


@pytest.mark.parametrize("name", ["sum", "count", "min", "max", "mean"])
def test_sf_range(sf: SheetFrame, df: DataFrame, name: str):
    func = Range((20, 1), sheet=sf.sheet)
    a = sf.agg(func, formula=True)
    b = df.agg(name)
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a.to_frame(), sheet=sf.sheet)
    func.value = name
    np.testing.assert_array_equal(sf.value[0], b)


def test_sf_str_columns(sf: SheetFrame):
    a = sf.agg("mean", columns="a", formula=True)
    assert len(a) == 1
    sf = SheetFrame(20, 2, data=a.to_frame(), sheet=sf.sheet)
    np.testing.assert_array_equal(sf.value, [[4.5]])


def test_sf_str_columns_list(sf: SheetFrame):
    a = sf.agg("mean", columns=["a", "b"], formula=True)
    assert len(a) == 2
    sf = SheetFrame(20, 2, data=a.to_frame(), sheet=sf.sheet)
    df = DataFrame([[4.5], [14.5]], index=["a", "b"], columns=["mean"])
    df_sf = sf.value
    assert df_sf.equals(df)
    assert df_sf.index.equals(df.index)
    assert df_sf.columns.equals(df.columns)


def test_sf_dict(sf: SheetFrame, df: DataFrame):
    func = {"a": "min", "b": "max"}
    a = sf.agg(func, formula=True)
    b = df.agg(func)
    assert isinstance(a, Series)
    assert a.index.to_list() == b.index.to_list()
    sf = SheetFrame(20, 2, data=a.to_frame(), sheet=sf.sheet)
    np.testing.assert_array_equal(sf.value[0], b)


def test_sf_list(sf: SheetFrame, df: DataFrame):
    func = ["min", "max"]
    a = sf.agg(func, formula=True)
    b = df.agg(func)  # type: ignore
    assert isinstance(a, DataFrame)
    assert a.index.to_list() == b.index.to_list()
    assert a.columns.to_list() == b.columns.to_list()
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet)
    np.testing.assert_array_equal(sf.value, b)


def test_sf_list_columns(sf: SheetFrame, df: DataFrame):
    a = sf.agg(["sum", "count"], columns="b", formula=True)
    assert isinstance(a, DataFrame)
    sf = SheetFrame(20, 2, data=a, sheet=sf.sheet)
    np.testing.assert_array_equal(sf.value, [[116], [8]])


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


def test_df_group_str_str(df: DataFrame):
    a = df.groupby("x").agg("sum")
    assert a.columns.to_list() == ["a", "b"]
    assert a.index.to_list() == [1, 2]
    assert a.index.name == "x"
    np.testing.assert_array_equal(a, [[10, 50], [26, 66]])


@pytest.mark.filterwarnings("ignore::FutureWarning")
def test_df_group_str_str_as_index_false(df: DataFrame):
    a = df.groupby("x", as_index=False).agg("sum")
    assert a.columns.to_list() == ["a", "b"]
    assert a.index.to_list() == [0, 1]
    assert a.index.name is None
    np.testing.assert_array_equal(a, [[10, 50], [26, 66]])


def test_df_group_list_str(df: DataFrame):
    a = df.groupby(["x", "y"]).agg("sum")
    assert a.columns.to_list() == ["a", "b"]
    assert a.index.to_list() == [(1, 1), (1, 2), (2, 1), (2, 2)]
    assert a.index.names == ["x", "y"]
    np.testing.assert_array_equal(a, [[3, 23], [7, 27], [11, 31], [15, 35]])


def test_df_group_list_list(df: DataFrame):
    a = df.groupby(["x", "y"]).agg(["min", "max"])
    c = [("a", "min"), ("a", "max"), ("b", "min"), ("b", "max")]
    assert a.columns.to_list() == c
    assert a.index.to_list() == [(1, 1), (1, 2), (2, 1), (2, 2)]
    assert a.index.names == ["x", "y"]
    v = [[1, 2, 11, 12], [3, 4, 13, 14], [5, 6, 15, 16], [7, 8, 17, 18]]
    np.testing.assert_array_equal(a, v)


def test_df_group_list_dict(df: DataFrame):
    a = df.groupby(["x", "y"]).agg({"a": "min", "b": "max"})
    assert a.columns.to_list() == ["a", "b"]
    assert a.index.to_list() == [(1, 1), (1, 2), (2, 1), (2, 2)]
    assert a.index.names == ["x", "y"]
    np.testing.assert_array_equal(a, [[1, 12], [3, 14], [5, 16], [7, 18]])


def test_index_str(sf: SheetFrame):
    a = sf.groupby("x").index()
    b = Index([1, 2], name="x")
    assert a.equals(b)


def test_index_list(sf: SheetFrame):
    a = sf.groupby(["x", "y"]).index()
    b = MultiIndex.from_tuples([(1, 1), (1, 2), (2, 1), (2, 2)], names=["x", "y"])
    assert a.equals(b)


def test_index_str_as_address(sf: SheetFrame):
    a = sf.groupby("x").index(as_address=True)
    b = Index(["$B$3", "$B$7"], name="x")
    assert a.equals(b)


def test_index_list_as_address(sf: SheetFrame):
    a = sf.groupby(["x", "y"]).index(as_address=True, formula=True)
    values = [(f"=$B${r}", f"=$C${r}") for r in [3, 5, 7, 9]]
    b = MultiIndex.from_tuples(values, names=["x", "y"])
    assert a.equals(b)


@pytest.mark.parametrize("func", ["sum", "median", "mean"])
@pytest.mark.parametrize("by", ["x", "y"])
def test_sf_group_str_str(sf: SheetFrame, df: DataFrame, func, by):
    a = sf.groupby(by).agg(func, as_address=True, formula=True)
    b = df.groupby(by).agg(func).astype(float)
    sf = SheetFrame(50, 2, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)


@pytest.mark.parametrize("func", ["sum", "median", "mean"])
@pytest.mark.parametrize("by", ["x", "y"])
def test_sf_group_str_range(sf: SheetFrame, df: DataFrame, func, by):
    rng = Range((50, 1), sheet=sf.sheet)
    rng.value = func
    a = sf.groupby(by).agg(rng, as_address=True, formula=True)
    b = df.groupby(by).agg(func).astype(float)
    sf = SheetFrame(50, 2, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)


@pytest.mark.parametrize("func", ["sum", "min", "max"])
@pytest.mark.parametrize("by", [["x", "y"], ["y", "x"]])
def test_sf_group_list_str(sf: SheetFrame, df: DataFrame, func, by):
    a = sf.groupby(by).agg(func, as_address=True, formula=True)
    b = df.groupby(by).agg(func).astype(float)
    sf = SheetFrame(50, 10, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)


@pytest.mark.parametrize("by", [["x", "y"], ["y", "x"]])
@pytest.mark.parametrize("sort", [True, False])
def test_sf_group_list_str_sort(sf: SheetFrame, df: DataFrame, by, sort):
    a = sf.groupby(by, sort=sort).agg("sum", as_address=True, formula=True)
    b = df.groupby(by, sort=sort).agg("sum").astype(float)
    sf = SheetFrame(50, 20, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)


@pytest.mark.parametrize(
    "func",
    [{"a": "min", "b": "max"}, {"a": "median", "b": "mean"}],
)
@pytest.mark.parametrize("by", [["x", "y"], ["y", "x"]])
@pytest.mark.parametrize("sort", [True, False])
def test_sf_group_list_dict(sf: SheetFrame, df: DataFrame, func, by, sort):
    a = sf.groupby(by, sort=sort).agg(func, as_address=True, formula=True)
    b = df.groupby(by, sort=sort).agg(func).astype(float)
    sf = SheetFrame(50, 30, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)


@pytest.mark.parametrize("func", [["sum", "mean"], ["min", "max"]])
@pytest.mark.parametrize("sort", [True, False])
def test_sf_group_list_list(sf: SheetFrame, df: DataFrame, func, sort):
    by = ["y", "x"]
    a = sf.groupby(by, sort=sort).agg(func, as_address=True, formula=True)
    b = df.groupby(by, sort=sort).agg(func).astype(float)
    sf = SheetFrame(50, 40, data=a, sheet=sf.sheet)
    df_sf = sf.value
    assert df_sf.equals(b)
    assert df_sf.index.equals(b.index)
    assert df_sf.columns.equals(b.columns)
