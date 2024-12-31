import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 2, 3, data=df, style=False)


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 2
    assert sf.column == 3
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 1
    assert sf.columns_level == 1


def test_value(sf: SheetFrame):
    v = [[None, "a", "b"], [0, 1, 5], [1, 2, 6], [2, 3, 7], [3, 4, 8]]
    assert sf.cell.expand().options(ndim=2).value == v


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns == [None, "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == [None]


def test_init_index_false(df: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df, index=False, style=False)
    assert sf.columns == ["a", "b"]
    assert sf.index_level == 0


def test_contains(sf: SheetFrame):
    assert "a" in sf
    assert "x" not in sf


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("a", True, 2),
        ("a", False, 4),
        ("b", True, 3),
        ("b", False, 5),
        (["a", "b"], True, [2, 3]),
        (["a", "b"], False, [4, 5]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_, df)


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$C$2:$E$6"
    assert sf.range().get_address() == "$C$2:$E$6"


def test_range_column_none(sf: SheetFrame):
    assert sf.range(column=None).get_address() == "$C$2:$E$6"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (None, None, "$C$3"),
        (False, None, "$C$2:$C$6"),
        (0, None, "$C$2"),
        (-1, None, "$C$3:$C$6"),
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
        ("a", 0, None, "$D$2"),
        ("b", 0, None, "$E$2"),
        ("a", 1, None, "$D$1"),
        ("a", 2, None, "$D$2"),
        ("b", 100, None, "$E$100"),
        ("a", -1, None, "$D$3:$D$6"),
        ("a", False, None, "$D$2:$D$6"),
        ("b", False, None, "$E$2:$E$6"),
        ("a", 2, 100, "$D$2:$D$100"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


def test_repr(sf: SheetFrame):
    assert repr(sf).endswith("!$C$2:$E$6>")


def test_str(sf: SheetFrame):
    assert str(sf).endswith("!$C$2:$E$6>")


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


def test_getitem_slice_none(sf: SheetFrame):
    df = sf[:]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == ["index", "a", "b"]
    x = [[0, 1, 5], [1, 2, 6], [2, 3, 7], [3, 4, 8]]
    np.testing.assert_array_equal(df, x)


def test_setitem(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(sheet, 2, 2, data=df, style=False)
    x = [10, 20, 30]
    sf["a"] = x
    np.testing.assert_array_equal(sf["a"], x)


def test_setitem_new_column(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(sheet, 2, 2, data=df, style=False)
    x = [10, 20, 30]
    sf["c"] = x
    assert sf.columns == [None, "a", "b", "c"]
    np.testing.assert_array_equal(sf["c"], x)


@pytest.mark.parametrize(
    ("a", "b", "sel"),
    [
        (1, None, [True, False, False, False]),
        (3, None, [False, False, True, False]),
        ([2, 4], None, [False, True, False, True]),
        ((2, 4), None, [False, True, True, True]),
        (1, 5, [True, False, False, False]),
        (1, 6, [False, False, False, False]),
        ((1, 3), (6, 8), [False, True, True, False]),
    ],
)
def test_select(sf: SheetFrame, a, b, sel):
    if b is None:
        np.testing.assert_array_equal(sf.select(a=a), sel)
    else:
        np.testing.assert_array_equal(sf.select(a=a, b=b), sel)


def test_groupby(sheet: Sheet):
    df = DataFrame({"a": [1, 1, 1, 2, 2, 1, 1], "b": [1, 2, 3, 4, 5, 6, 7]})
    sf = SheetFrame(sheet, 2, 2, data=df, style=False, index=False)

    g = sf.groupby("a")
    assert g[1.0] == [[3, 5], [8, 9]]
    assert g[2.0] == [[6, 7]]

    assert len(sf.groupby(["a", "b"])) == 7

    g = sf.groupby("::b")
    assert g[(1.0,)] == [[3, 5], [8, 9]]
    assert g[(2.0,)] == [[6, 7]]

    assert len(sf.groupby(":b")) == 7


def test_row_one(sheet: Sheet):
    df = DataFrame({"a": [1], "b": [2]})
    sf = SheetFrame(sheet, 2, 2, data=df, style=False)
    assert len(sf) == 1
    np.testing.assert_array_equal(sf["a"], [1])


def test_column_one(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3]})
    sf = SheetFrame(sheet, 2, 2, data=df, style=False, index=False)
    assert len(sf) == 3
    assert sf.columns == ["a"]
    np.testing.assert_array_equal(sf["a"], [1, 2, 3])
