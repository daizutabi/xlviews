import numpy as np
import pytest
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 2, 2, 2, 2],
            "y": [1, 1, 2, 2, 1, 1, 2, 2],
            "z": [1, 1, 1, 2, 2, 1, 1, 1],
            "c": [1, 2, 3, 4, 5, 6, 7, 8],
            "d": [11, 12, 13, 14, 15, 16, 17, 18],
            "e": [21, 22, 23, 24, 25, 26, 27, 28],
            "f": [31, 32, 33, 34, 35, 36, 37, 38],
        },
    )
    df = df.set_index(["x", "y", "z"])
    x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    df.columns = MultiIndex.from_tuples(x, names=["a", "b"])
    return df


def test_df(df: DataFrame):
    assert len(df) == 8
    assert df.shape == (8, 4)

    x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert df.columns.to_list() == x
    assert df.columns.names == ["a", "b"]
    assert isinstance(df.columns, MultiIndex)

    x = [(1, 1, 1), (1, 1, 1), (1, 2, 1), (1, 2, 2), (2, 1, 2), (2, 1, 1)]
    x += [(2, 2, 1), (2, 2, 1)]
    assert df.index.to_list() == x
    assert df.index.names == ["x", "y", "z"]
    assert isinstance(df.index, MultiIndex)


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 5, 3, data=df, style=False)


def test_value_column(sf: SheetFrame):
    assert sf.cell.value is None
    v = sf.cell.offset(0, 3).expand().options(ndim=2).value
    assert len(v) == 10
    assert v[0] == ["a1", "a1", "a2", "a2"]
    assert v[1] == ["b1", "b2", "b1", "b2"]
    assert v[2] == [1, 11, 21, 31]
    assert v[-1] == [8, 18, 28, 38]


def test_value_values(sf: SheetFrame):
    v = sf.cell.offset(1, 0).expand().options(ndim=2).value
    assert len(v) == 9
    assert v[0] == ["x", "y", "z", "b1", "b2", "b1", "b2"]
    assert v[1] == [1, 1, 1, 1, 11, 21, 31]
    assert v[2] == [1, 1, 1, 2, 12, 22, 32]
    assert v[-1] == [2, 2, 1, 8, 18, 28, 38]


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$C$5"
    assert sf.row == 5
    assert sf.column == 3
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 3
    assert sf.columns_level == 2
    assert sf.columns_names is None


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_columns(sf: SheetFrame):
    i = "x", "y", "z"
    c = ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")
    assert sf.columns == [*i, *c]


def test_value_columns(sf: SheetFrame):
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y", "z"]


def test_init_index_false(df: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df, index=False, style=False)

    assert sf.index_level == 0
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.columns == c


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "z" in sf
    assert ("a1", "b1") in sf
    assert "a1" not in sf


def test_iter(sf: SheetFrame):
    i = "x", "y", "z"
    c = ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")
    assert list(sf) == [*i, *c]


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("x", True, 1),
        ("z", False, 5),
        (("a1", "b1"), True, 4),
        (("a1", "b1"), False, 6),
        (["y", ("a2", "b1")], True, [2, 6]),
        (["x", ("a2", "b2")], False, [3, 9]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


def test_index_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="'a' is not in list"):
        sf.index("a")


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    # np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name
    # assert df_.columns.name == df.columns.name


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$C$5:$I$14"
    assert sf.range().get_address() == "$C$5:$I$14"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (0, None, "$C$5:$E$6"),
        (None, None, "$C$7:$E$7"),
        (-1, None, "$C$7:$E$14"),
        (False, None, "$C$5:$E$14"),
        (20, None, "$C$20:$E$20"),
        (20, 100, "$C$20:$E$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


# @pytest.mark.parametrize(
#     ("column", "start", "end", "address"),
#     [
#         ("x", 0, None, "$F$10"),
#         ("y", 0, None, "$G$10"),
#         ("a", 0, None, "$H$10"),
#         ("b", 0, None, "$I$10"),
#         ("x", 1, None, "$F$1"),
#         ("a", 2, None, "$H$2"),
#         ("b", 100, None, "$I$100"),
#         ("y", -1, None, "$G$11:$G$18"),
#         ("a", False, None, "$H$10:$H$18"),
#         ("b", False, None, "$I$10:$I$18"),
#         ("x", 2, 100, "$F$2:$F$100"),
#         ("a", 3, 300, "$H$3:$H$300"),
#     ],
# )
# def test_range_column(sf: SheetFrame, column, start, end, address):
#     assert sf.range_column(column, start, end).get_address() == address
#     assert sf.range(column, start, end).get_address() == address


# def test_repr(sf: SheetFrame):
#     assert repr(sf).endswith("!$C$2:$E$6>")


# def test_str(sf: SheetFrame):
#     assert str(sf).endswith("!$C$2:$E$6>")


# def test_getitem_str(sf: SheetFrame):
#     s = sf["a"]
#     assert isinstance(s, Series)
#     assert s.name == "a"
#     np.testing.assert_array_equal(s, [1, 2, 3, 4])


# def test_getitem_list(sf: SheetFrame):
#     df = sf[["a", "b"]]
#     assert isinstance(df, DataFrame)
#     assert df.columns.to_list() == ["a", "b"]
#     x = [[1, 5], [2, 6], [3, 7], [4, 8]]
#     np.testing.assert_array_equal(df, x)


# def test_getitem_slice_none(sf: SheetFrame):
#     df = sf[:]
#     assert isinstance(df, DataFrame)
#     assert df.columns.to_list() == ["index", "a", "b"]
#     x = [[0, 1, 5], [1, 2, 6], [2, 3, 7], [3, 4, 8]]
#     np.testing.assert_array_equal(df, x)


# def test_setitem(sheet: Sheet):
#     df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False)
#     x = [10, 20, 30]
#     sf["a"] = x
#     np.testing.assert_array_equal(sf["a"], x)


# def test_setitem_new_column(sheet: Sheet):
#     df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False)
#     x = [10, 20, 30]
#     sf["c"] = x
#     assert sf.columns == [None, "a", "b", "c"]
#     np.testing.assert_array_equal(sf["c"], x)


# @pytest.mark.parametrize(
#     ("a", "b", "sel"),
#     [
#         (1, None, [True, False, False, False]),
#         (3, None, [False, False, True, False]),
#         ([2, 4], None, [False, True, False, True]),
#         ((2, 4), None, [False, True, True, True]),
#         (1, 5, [True, False, False, False]),
#         (1, 6, [False, False, False, False]),
#         ((1, 3), (6, 8), [False, True, True, False]),
#     ],
# )
# def test_select(sf: SheetFrame, a, b, sel):
#     if b is None:
#         np.testing.assert_array_equal(sf.select(a=a), sel)
#     else:
#         np.testing.assert_array_equal(sf.select(a=a, b=b), sel)


# def test_groupby(sheet: Sheet):
#     df = DataFrame({"a": [1, 1, 1, 2, 2, 1, 1], "b": [1, 2, 3, 4, 5, 6, 7]})
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False, index=False)

#     g = sf.groupby("a")
#     assert g[1.0] == [[3, 5], [8, 9]]
#     assert g[2.0] == [[6, 7]]

#     assert len(sf.groupby(["a", "b"])) == 7

#     g = sf.groupby("::b")
#     assert g[(1.0,)] == [[3, 5], [8, 9]]
#     assert g[(2.0,)] == [[6, 7]]

#     assert len(sf.groupby(":b")) == 7


# def test_row_one(sheet: Sheet):
#     df = DataFrame({"a": [1], "b": [2]})
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False)
#     assert len(sf) == 1
#     np.testing.assert_array_equal(sf["a"], [1])


# def test_column_one(sheet: Sheet):
#     df = DataFrame({"a": [1, 2, 3]})
#     sf = SheetFrame(sheet, 2, 2, data=df, style=False, index=False)
#     assert len(sf) == 3
#     assert sf.columns == ["a"]
#     np.testing.assert_array_equal(sf["a"], [1, 2, 3])
