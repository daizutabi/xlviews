import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame
from xlviews.table import Table


@pytest.fixture(scope="module")
def df():
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]}, index=["x", "x", "y", "y"])
    df.index.name = "name"
    return df


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(2, 3, data=df, style=False, sheet=sheet_module)


def test_value(sf: SheetFrame):
    v = [["name", "a", "b"], ["x", 1, 5], ["x", 2, 6], ["y", 3, 7], ["y", 4, 8]]
    assert sf.cell.expand().options(ndim=2).value == v


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 2
    assert sf.column == 3
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 1
    assert sf.columns_level == 1


def test_set_data_from_sheet(sf: SheetFrame):
    sf.set_data_from_sheet(index_level=0)
    assert sf.has_index is False
    assert sf.value_columns == ["name", "a", "b"]
    sf.set_data_from_sheet(index_level=1)
    assert sf.has_index is True
    assert sf.value_columns == ["a", "b"]


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


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("name", True, 1),
        ("name", False, 3),
        ("a", True, 2),
        ("a", False, 4),
        ("b", True, 3),
        ("b", False, 5),
        (["name", "b"], True, [1, 3]),
        (["name", "b"], False, [3, 5]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


@pytest.fixture
def table(sf: SheetFrame):
    yield sf.as_table()
    sf.unlist()


@pytest.mark.parametrize("value", ["x", "y"])
def test_table(table: Table, value):
    table.auto_filter("name", value)
    header = table.const_header.value
    assert isinstance(header, list)
    assert header[0] == value


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("x", [[1, 5], [2, 6]]),
        ("y", [[3, 7], [4, 8]]),
    ],
)
def test_visible_data(sf: SheetFrame, table: Table, name, value):
    table.auto_filter("name", name)
    df = sf.visible_data
    assert df.index.to_list() == [name, name]
    np.testing.assert_array_equal(df, value)


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$C$2:$E$6"
    assert sf.range().get_address() == "$C$2:$E$6"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (False, None, "$C$2:$C$6"),
        (-1, None, "$C$2"),
        (0, None, "$C$3"),
        (None, None, "$C$3:$C$6"),
        (10, None, "$C$10"),
        (10, 100, "$C$10:$C$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        ("a", -1, None, "$D$2"),
        ("b", -1, None, "$E$2"),
        ("name", -1, None, "$C$2"),
        ("a", 1, None, "$D$1"),
        ("a", 2, None, "$D$2"),
        ("b", 100, None, "$E$100"),
        ("a", None, None, "$D$3:$D$6"),
        ("a", False, None, "$D$2:$D$6"),
        ("b", False, None, "$E$2:$E$6"),
        ("name", False, None, "$C$2:$C$6"),
        ("a", 2, 100, "$D$2:$D$100"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "address"),
    [("a", "$D$3"), ("b", "$E$3"), ("name", "$C$3")],
)
def test_first_range(sf: SheetFrame, column, address):
    assert sf.first_range(column).get_address() == address


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


@pytest.mark.parametrize(
    ("name", "a", "sel"),
    [
        ("x", None, [True, True, False, False]),
        ("y", None, [False, False, True, True]),
        ("x", (2, 3), [False, True, False, False]),
        ("x", [1, 4], [True, False, False, False]),
        (["x", "y"], [1, 4], [True, False, False, True]),
    ],
)
def test_select(sf: SheetFrame, name, a, sel):
    if a is None:
        np.testing.assert_array_equal(sf.select(name=name), sel)
    else:
        np.testing.assert_array_equal(sf.select(name=name, a=a), sel)


def test_groupby(sf: SheetFrame):
    g = sf.groupby("name")
    assert len(g) == 2
    assert g[("x",)] == [(3, 4)]
    assert g[("y",)] == [(5, 6)]

    assert len(sf.groupby(["name", "a"])) == 4


@pytest.fixture(scope="module")
def sf2(sheet_module: Sheet):
    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    return SheetFrame(102, 2, data=df, index=True, style=False, sheet=sheet_module)


@pytest.mark.parametrize(
    ("kwargs", "r"),
    [({}, range(103, 113)), ({"c": 100}, [103, 104, 108, 109])],
)
def test_ranges(sf2: SheetFrame, kwargs, r):
    for rng, i in zip(sf2.ranges(**kwargs), r, strict=True):
        assert rng.get_address() == f"$E${i}:$F${i}"


def test_ranges_sel(sf2: SheetFrame):
    sel = sf2.select(c=200)
    it = sf2.ranges(sel, b="t")

    for rng, i in zip(it, [110, 111, 112], strict=True):
        assert rng.get_address() == f"$E${i}:$F${i}"
