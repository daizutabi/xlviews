import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.table import Table
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame import Index

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Index(sheet_module, 2, 3)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame):
    assert sf.row == 2
    assert sf.column == 3
    assert sf.index_level == 1
    assert sf.columns_level == 1
    assert sf.columns_names is None


def test_set_data_from_sheet(sf: SheetFrame):
    sf.set_data_from_sheet(index_level=0)
    assert sf.has_index is False
    assert sf.value_columns == ["name", "a", "b"]
    sf.set_data_from_sheet(index_level=1)
    assert sf.has_index is True
    assert sf.value_columns == ["a", "b"]


def test_expand(sf: SheetFrame):
    v = [["name", "a", "b"], ["x", 1, 5], ["x", 2, 6], ["y", 3, 7], ["y", 4, 8]]
    assert sf.expand().options(ndim=2).value == v


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns == ["name", "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["name"]


def test_contains(sf: SheetFrame):
    assert "name" in sf
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["name", "a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("name", 3),
        ("a", 4),
        ("b", 5),
        (["name", "b"], [3, 5]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index(column) == index


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("b", 0, "$E$3"),
        ("a", -1, "$D$2"),
        ("b", -1, "$E$2"),
        ("name", -1, "$C$2"),
        ("a", None, "$D$3:$D$6"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


def test_getitem_str(sf: SheetFrame):
    s = sf["a"]
    assert isinstance(s, Series)
    assert s.name == "a"
    np.testing.assert_array_equal(s, [1, 2, 3, 4])


def test_getitem_list(sf: SheetFrame):
    df = sf[["a", "b"]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == ["a", "b"]
    x = [[1, 5], [2, 6], [3, 7], [4, 8]]
    np.testing.assert_array_equal(df, x)


# @pytest.mark.parametrize(
#     ("column", "address"),
#     [("a", "$D$3"), ("b", "$E$3"), ("name", "$C$3")],
# )
# def test_first_range(sf: SheetFrame, column, address):
#     assert sf.first_range(column).get_address() == address


# def test_groupby(sf: SheetFrame):
#     g = groupby(sf, "name")
#     assert len(g) == 2
#     assert g[("x",)] == [(3, 4)]
#     assert g[("y",)] == [(5, 6)]

#     assert len(groupby(sf, ["name", "a"])) == 4


# @pytest.fixture(scope="module")
# def sf2(sheet_module: Sheet):
#     a = ["c"] * 10
#     b = ["s"] * 5 + ["t"] * 5
#     c = ([100] * 2 + [200] * 3) * 2
#     x = list(range(10))
#     y = list(range(10, 20))
#     df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
#     df = df.set_index(["a", "b", "c"])
#     return SheetFrame(102, 2, data=df, index=True, style=False, sheet=sheet_module)


# @pytest.mark.parametrize(
#     ("kwargs", "r"),
#     [({}, range(103, 113)), ({"c": 100}, [103, 104, 108, 109])],
# )
# def test_ranges(sf2: SheetFrame, kwargs, r):
#     for rng, i in zip(sf2.ranges(**kwargs), r, strict=True):
#         assert rng.get_address() == f"$E${i}:$F${i}"


# def test_ranges_sel(sf2: SheetFrame):
#     sel = sf2.select(c=200)
#     it = sf2.ranges(sel, b="t")

#     for rng, i in zip(it, [110, 111, 112], strict=True):
#         assert rng.get_address() == f"$E${i}:$F${i}"


# @pytest.fixture(scope="module")
# def address(sf: SheetFrame):
#     return sf.get_address()


# def test_get_address_index_name(address: DataFrame):
#     assert address.index.name == "name"


# def test_get_address_index(address: DataFrame):
#     assert address.index.to_list() == ["x", "x", "y", "y"]


# def test_get_address_value(address: DataFrame):
#     values = [["$D$3", "$E$3"], ["$D$4", "$E$4"], ["$D$5", "$E$5"], ["$D$6", "$E$6"]]
#     np.testing.assert_array_equal(address, values)
