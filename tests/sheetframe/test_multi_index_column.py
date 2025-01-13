import numpy as np
import pytest
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


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
    return SheetFrame(5, 3, data=df, style=False, sheet=sheet_module)


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


def test_set_data_from_sheet(sf: SheetFrame):
    sf.set_data_from_sheet(index_level=2, columns_level=2)
    assert sf.has_index is True
    assert sf.index_columns == ["x", "y"]
    c = [(None, "z"), ("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c
    sf.set_data_from_sheet(index_level=3, columns_level=2)
    assert sf.has_index is True
    assert sf.index_columns == ["x", "y", "z"]
    c = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == c


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
    sf = SheetFrame(2, 3, data=df, index=False, style=False, sheet=sheet)

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
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$C$5:$I$14"
    assert sf.range().get_address() == "$C$5:$I$14"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (False, None, "$C$5:$E$14"),
        (-1, None, "$C$5:$E$6"),
        (0, None, "$C$7:$E$7"),
        (None, None, "$C$7:$E$14"),
        (20, None, "$C$20:$E$20"),
        (20, 100, "$C$20:$E$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        ("x", -1, None, "$C$5:$C$6"),
        ("y", 0, None, "$D$7"),
        ("z", None, None, "$E$7:$E$14"),
        (("a2", "b2"), False, None, "$I$5:$I$14"),
        (("a1", "b1"), -1, None, "$F$5:$F$6"),
        (("a1", "b2"), 0, None, "$G$7"),
        (("a2", "b1"), None, None, "$H$7:$H$14"),
        ("y", 30, None, "$D$30"),
        (("a1", "b2"), 30, 40, "$G$30:$G$40"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


def test_getitem_str(sf: SheetFrame):
    s = sf["x"]
    assert isinstance(s, Series)
    assert s.name == "x"
    np.testing.assert_array_equal(s, [1, 1, 1, 1, 2, 2, 2, 2])


def test_getitem_list(sf: SheetFrame):
    df = sf[["z", ("a1", "b1")]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == ["z", ("a1", "b1")]
    x = np.array([[1, 1, 1, 2, 2, 1, 1, 1], range(1, 9)]).T
    np.testing.assert_array_equal(df, x)


@pytest.mark.parametrize(
    ("by", "one", "two"),
    [
        ("x", [(7, 10)], [(11, 14)]),
        ("y", [(7, 8), (11, 12)], [(9, 10), (13, 14)]),
        ("z", [(7, 9), (12, 14)], [(10, 11)]),
    ],
)
def test_groupby(sf: SheetFrame, by, one, two):
    g = sf.groupby(by)
    assert len(g) == 2
    assert g[(1,)] == one
    assert g[(2,)] == two


@pytest.mark.parametrize(
    ("by", "v11", "v12", "v21", "v22"),
    [
        (["x", "y"], [(7, 8)], [(9, 10)], [(11, 12)], [(13, 14)]),
        (["x", "z"], [(7, 9)], [(10, 10)], [(12, 14)], [(11, 11)]),
        (["y", "z"], [(7, 8), (12, 12)], [(11, 11)], [(9, 9), (13, 14)], [(10, 10)]),
    ],
)
def test_groupby_list(sf: SheetFrame, by, v11, v12, v21, v22):
    g = sf.groupby(by)
    assert len(g) == 4
    assert g[(1, 1)] == v11
    assert g[(1, 2)] == v12
    assert g[(2, 1)] == v21
    assert g[(2, 2)] == v22
