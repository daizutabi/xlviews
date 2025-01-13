import numpy as np
import pytest
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


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
    return SheetFrame(20, 4, data=df, style=False, sheet=sheet_module)


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


def test_set_data_from_sheet(sf: SheetFrame):
    sf.set_data_from_sheet(index_level=1, columns_level=2)
    assert sf.has_index is True
    x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
    assert sf.value_columns == x
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
    sf = SheetFrame(2, 3, data=df, index=False, style=False, sheet=sheet)

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
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns, df.columns)
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_, df)
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$D$20:$H$26"
    assert sf.range().get_address() == "$D$20:$H$26"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (0, None, "$D$22"),
        (50, None, "$D$50"),
        (50, 100, "$D$50:$D$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize("start", [None, -1, False])
def test_range_index_error(sf: SheetFrame, start):
    with pytest.raises(ValueError, match="index start must be a specific row"):
        sf.range_index(start)

    with pytest.raises(ValueError, match="index start must be a specific row"):
        sf.range("index", start)


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        (("a2", "b2"), False, None, "$H$20:$H$26"),
        (("a1", "b1"), -1, None, "$E$20:$E$21"),
        (("a1", "b2"), 0, None, "$F$22"),
        (("a2", "b1"), None, None, "$G$22:$G$26"),
        (("a1", "b1"), 30, None, "$E$30"),
        (("a1", "b2"), 30, 40, "$F$30:$F$40"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


def test_getitem_tuple(sf: SheetFrame):
    s = sf[("a1", "b1")]
    assert isinstance(s, Series)
    assert s.name == ("a1", "b1")
    np.testing.assert_array_equal(s, [1, 2, 3, 4, 5])


def test_getitem_list(sf: SheetFrame):
    df = sf[[("a1", "b1"), ("a2", "b2")]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == [("a1", "b1"), ("a2", "b2")]
    x = np.array([[1, 2, 3, 4, 5], [31, 32, 33, 34, 35]]).T
    np.testing.assert_array_equal(df, x)


@pytest.mark.parametrize(
    ("a", "b", "sel"),
    [
        ("a1", None, [True, True, False, False]),
        ("a2", None, [False, False, True, True]),
        ("a1", "b1", [True, False, False, False]),
        ("a2", "b2", [False, False, False, True]),
        (["a1", "a2"], None, [True, True, True, True]),
        (["a1", "a2"], ["b2"], [False, True, False, True]),
        (("a1", "a2"), ("b1", "b1"), [True, False, True, False]),
    ],
)
def test_select(sf: SheetFrame, a, b, sel):
    if b is None:
        np.testing.assert_array_equal(sf.select(a=a), sel)
    else:
        np.testing.assert_array_equal(sf.select(a=a, b=b), sel)


@pytest.mark.parametrize(
    ("by", "one", "two"),
    [
        ("a", [(5, 6)], [(7, 8)]),
        ("b", [(5, 5), (7, 7)], [(6, 6), (8, 8)]),
    ],
)
def test_groupby(sf: SheetFrame, by, one, two):
    g = sf.groupby(by)
    assert len(g) == 2
    assert g[(f"{by}1",)] == one
    assert g[(f"{by}2",)] == two


def test_groupby_list(sf: SheetFrame):
    g = sf.groupby(["a", "b"])
    assert len(g) == 4
    assert g[("a1", "b1")] == [(5, 5)]
    assert g[("a1", "b2")] == [(6, 6)]
    assert g[("a2", "b1")] == [(7, 7)]
    assert g[("a2", "b2")] == [(8, 8)]


def test_groupby_none(sf: SheetFrame):
    g = sf.groupby(None)
    assert len(g) == 1
    assert g[()] == [(5, 8)]
