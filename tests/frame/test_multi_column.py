import numpy as np
import pytest
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "c": [1, 2, 3, 4, 5],
            "d": [11, 12, 13, 14, 15],
            "e": [21, 22, 23, 24, 25],
            "f": [31, 32, 33, 34, 35],
        },
    )
    x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    df.columns = MultiIndex.from_tuples(x, names=["a", "b"])
    return df


def test_df(df: DataFrame):
    assert len(df) == 5
    assert df.shape == (5, 4)

    x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert df.columns.to_list() == x
    assert df.columns.names == ["a", "b"]
    assert isinstance(df.columns, MultiIndex)

    assert df.index.to_list() == list(range(5))
    assert df.index.names == [None]


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 20, 4, data=df, style=False)


def test_value(sf: SheetFrame):
    v = sf.cell.expand().options(ndim=2).value
    assert len(v) == 7
    assert v[0] == ["a", "a1", "a1", "a2", "a2"]
    assert v[1] == ["b", "b1", "b2", "b1", "b2"]
    assert v[2] == [0, 1, 11, 21, 31]
    assert v[-1] == [4, 5, 15, 25, 35]


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$D$20"
    assert sf.row == 20
    assert sf.column == 4
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 1
    assert sf.columns_level == 2
    assert sf.columns_names == ["a", "b"]


def test_len(sf: SheetFrame):
    assert len(sf) == 5


def test_columns(sf: SheetFrame):
    x = [("a", "b"), ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.columns == x


def test_value_columns(sf: SheetFrame):
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == [("a", "b")]


def test_init_index_false(df: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df, index=False, style=False)

    assert sf.index_level == 0
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.columns == c


def test_contains(sf: SheetFrame):
    assert None not in sf
    assert ("a", "b") in sf
    assert ("a1", "b1") in sf
    assert "a1" not in sf


def test_iter(sf: SheetFrame):
    x = [("a", "b"), ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert list(sf) == x


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        (("a", "b"), True, 1),
        (("a1", "b1"), True, 2),
        (("a2", "b2"), True, 5),
        (("a1", "b2"), False, 6),
        (("a2", "b1"), False, 7),
        ([("a1", "b2"), ("a2", "b1")], True, [3, 4]),
        ([("a", "b"), ("a2", "b2")], False, [4, 8]),
        ("a", True, 1),
        (["b", "a"], True, [2, 1]),
        ("a", False, 20),
        (["a", "b"], False, [20, 21]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


@pytest.mark.parametrize("column", ["a", "b"])
def test_index_row(sf: SheetFrame, column):
    r = sf.index(column, relative=False)
    assert sf.sheet.range(r, sf.column).value == column


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ({"a": "a1"}, True, (2, 3)),
        ({"a": "a2"}, True, (4, 5)),
        ({"a": "a1"}, False, (5, 6)),
        ({"a": "a2"}, False, (7, 8)),
    ],
)
def test_index_dict(sf: SheetFrame, column, relative, index):
    assert sf.index_dict(column, relative=relative) == index
    assert sf.index(column, relative=relative) == index


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    print(df_)
    print(df_.index)
    print(df_.columns)
    assert 0
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_, df)


# def test_range_all(sf: SheetFrame):
#     assert sf.range_all().get_address() == "$F$10:$I$18"
#     assert sf.range().get_address() == "$F$10:$I$18"


# def test_range_column_none(sf: SheetFrame):
#     assert sf.range(column=None).get_address() == "$F$10:$I$18"


# @pytest.mark.parametrize(
#     ("start", "end", "address"),
#     [
#         (0, None, "$F$10:$G$10"),
#         (None, None, "$F$11:$G$11"),
#         (-1, None, "$F$11:$G$18"),
#         (False, None, "$F$10:$G$18"),
#         (20, None, "$F$20:$G$20"),
#         (20, 100, "$F$20:$G$100"),
#     ],
# )
# def test_range_index(sf: SheetFrame, start, end, address):
#     assert sf.range_index(start, end).get_address() == address
#     assert sf.range("index", start, end).get_address() == address


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
