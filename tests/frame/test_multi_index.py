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
            "a": [1, 2, 3, 4, 5, 6, 7, 8],
            "b": [11, 12, 13, 14, 15, 16, 17, 18],
        },
    )
    return df.set_index(["x", "y"])


def test_df(df: DataFrame):
    assert len(df) == 8
    assert df.shape == (8, 2)
    assert df.columns.to_list() == ["a", "b"]
    x = [(1, 1), (1, 1), (1, 2), (1, 2), (2, 1), (2, 1), (2, 2), (2, 2)]
    assert df.index.to_list() == x
    assert df.index.names == ["x", "y"]
    assert isinstance(df.index, MultiIndex)


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 10, 6, data=df, style=False)


def test_value(sf: SheetFrame):
    v = sf.cell.expand().options(ndim=2).value
    assert len(v) == 9
    assert v[0] == ["x", "y", "a", "b"]
    assert v[1] == [1, 1, 1, 11]
    assert v[2] == [1, 1, 2, 12]
    assert v[3] == [1, 2, 3, 13]
    assert v[4] == [1, 2, 4, 14]
    assert v[5] == [2, 1, 5, 15]
    assert v[6] == [2, 1, 6, 16]
    assert v[7] == [2, 2, 7, 17]
    assert v[8] == [2, 2, 8, 18]


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 10
    assert sf.column == 6
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 2
    assert sf.columns_level == 1


def test_init_index_false(df: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df, index=False, style=False)
    assert sf.columns == ["a", "b"]
    assert sf.index_level == 0


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_columns(sf: SheetFrame):
    assert sf.columns == ["x", "y", "a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y"]


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "a" in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["x", "y", "a", "b"]


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("a", True, 3),
        ("a", False, 8),
        ("b", True, 4),
        ("b", False, 9),
        (["x", "b"], True, [1, 4]),
        (["y", "b"], False, [7, 9]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


def test_index_error(sf: SheetFrame):
    with pytest.raises(IndexError):
        sf.index("z")


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_, df)


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$F$10:$I$18"
    assert sf.range().get_address() == "$F$10:$I$18"


def test_range_column_none(sf: SheetFrame):
    assert sf.range(column=None).get_address() == "$F$10:$I$18"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (0, None, "$F$10:$G$10"),
        (None, None, "$F$11:$G$11"),
        (-1, None, "$F$11:$G$18"),
        (False, None, "$F$10:$G$18"),
        (20, None, "$F$20:$G$20"),
        (20, 100, "$F$20:$G$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        ("x", 0, None, "$F$10"),
        ("y", 0, None, "$G$10"),
        ("a", 0, None, "$H$10"),
        ("b", 0, None, "$I$10"),
        ("x", 1, None, "$F$1"),
        ("a", 2, None, "$H$2"),
        ("b", 100, None, "$I$100"),
        ("y", -1, None, "$G$11:$G$18"),
        ("a", False, None, "$H$10:$H$18"),
        ("b", False, None, "$I$10:$I$18"),
        ("x", 2, 100, "$F$2:$F$100"),
        ("a", 3, 300, "$H$3:$H$300"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


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