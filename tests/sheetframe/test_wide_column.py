import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame({"x": ["i", "j"], "y": ["k", "l"], "a": [1, 2], "b": [3, 4]})
    return df.set_index(["x", "y"])


def test_df(df: DataFrame):
    assert len(df) == 2
    assert df.shape == (2, 2)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.to_list() == [("i", "k"), ("j", "l")]
    assert df.index.names == ["x", "y"]


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    sf = SheetFrame(4, 2, data=df, style=False, sheet=sheet_module)
    sf.add_wide_column("u", range(3), autofit=False)
    sf.add_wide_column("v", range(4), autofit=False)
    return sf


def test_value(sf: SheetFrame):
    v = sf.cell.expand().options(ndim=2).value
    assert len(v) == 3
    assert v[0] == ["x", "y", "a", "b", *range(3), *range(4)]
    assert v[1] == ["i", "k", 1, 3, *([None] * 7)]
    assert v[2] == ["j", "l", 2, 4, *([None] * 7)]

    assert sf.cell.offset(-1, 4).value == "u"
    assert sf.cell.offset(-1, 7).value == "v"


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 4
    assert sf.column == 2
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 2
    assert sf.columns_level == 1


def test_len(sf: SheetFrame):
    assert len(sf) == 2


def test_columns(sf: SheetFrame):
    assert sf.columns == ["x", "y", "a", "b", *range(3), *range(4)]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b", *range(3), *range(4)]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y"]


def test_wide_columns(sf: SheetFrame):
    assert sf.wide_columns == ["u", "v"]


def test_contains(sf: SheetFrame):
    assert "x" in sf
    assert "a" in sf
    assert "u" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["x", "y", "a", "b", *range(3), *range(4)]


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("a", True, 3),
        ("a", False, 4),
        ("b", True, 4),
        ("b", False, 5),
        (["x", "b"], True, [1, 4]),
        (["y", "b"], False, [3, 5]),
    ],
)
def test_index(sf: SheetFrame, column, relative, index):
    assert sf.index(column, relative=relative) == index


@pytest.mark.parametrize(
    ("column", "relative", "index"),
    [
        ("u", True, (5, 7)),
        ("v", True, (8, 11)),
        ("u", False, (6, 8)),
        ("v", False, (9, 12)),
        (("u", 0), True, 5),
        (("u", 2), True, 7),
        (("v", 0), False, 9),
        (("v", 3), False, 12),
    ],
)
def test_index_wide(sf: SheetFrame, column, relative, index):
    assert sf.index_wide(column, relative=relative) == index


@pytest.mark.parametrize("column", ["z", ("u", -1)])
def test_index_error(sf: SheetFrame, column):
    with pytest.raises(ValueError, match=".* is not in list"):
        sf.index(column)


def test_data(sf: SheetFrame, df: DataFrame):
    df_ = sf.data
    np.testing.assert_array_equal(df_.index, df.index)
    np.testing.assert_array_equal(df_.index.names, df.index.names)
    np.testing.assert_array_equal(df_.columns[:2], df.columns)
    np.testing.assert_array_equal(df_.columns[2:], [*range(3), *range(4)])
    np.testing.assert_array_equal(df_.columns.names, df.columns.names)
    np.testing.assert_array_equal(df_.iloc[:, :2], df)
    assert df_.iloc[:, 2:].isna().all().all()
    assert df_.index.name == df.index.name
    assert df_.columns.name == df.columns.name


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$B$4:$L$6"
    assert sf.range().get_address() == "$B$4:$L$6"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (False, None, "$B$4:$C$6"),
        (-1, None, "$B$4:$C$4"),
        (0, None, "$B$5:$C$5"),
        (None, None, "$B$5:$C$6"),
        (20, None, "$B$20:$C$20"),
        (20, 100, "$B$20:$C$100"),
    ],
)
def test_range_index(sf: SheetFrame, start, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        ("b", False, None, "$E$4:$E$6"),
        ("x", -1, None, "$B$4"),
        ("y", 0, None, "$C$5"),
        ("a", None, None, "$D$5:$D$6"),
        (("u", 0), -1, None, "$F$3:$F$4"),
        (("u", 2), 0, None, "$H$5"),
        (("v", 0), None, None, "$I$5:$I$6"),
        (("v", 3), False, None, "$L$3:$L$6"),
        ("u", -1, None, "$F$3:$H$4"),
        ("u", 0, None, "$F$5:$H$5"),
        ("v", None, None, "$I$5:$L$6"),
        ("v", False, None, "$I$3:$L$6"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "value"),
    [("a", [1, 2]), (("u", 1), [np.nan, np.nan])],
)
def test_getitem_str(sf: SheetFrame, column, value):
    s = sf[column]
    assert isinstance(s, Series)
    assert s.name == column
    np.testing.assert_array_equal(s, value)


def test_getitem_list(sf: SheetFrame):
    df = sf[["b", ("v", 3)]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == ["b", ("v", 3)]
    x = [[3, np.nan], [4, np.nan]]
    np.testing.assert_array_equal(df, x)
