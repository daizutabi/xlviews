import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_excel_installed
from xlviews.testing.sheet_frame import NoIndex

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return NoIndex(sheet_module, 2, 3)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, fc: FrameContainer):
    assert sf.row == fc.row
    assert sf.column == fc.column
    assert sf.index_level == 0
    assert sf.columns_level == 1
    assert sf.columns_names is None


def test_set_data_from_sheet(sf: SheetFrame):
    sf.set_data_from_sheet(index_level=1)
    assert sf.has_index is True
    assert sf.value_columns == ["b"]
    sf.set_data_from_sheet(index_level=0)
    assert sf.has_index is False
    assert sf.value_columns == ["a", "b"]


def test_expand(sf: SheetFrame):
    v = [["a", "b"], [1, 5], [2, 6], [3, 7], [4, 8]]
    assert sf.expand().options(ndim=2).value == v


def test_repr(sf: SheetFrame):
    assert repr(sf).endswith("!$C$2:$D$6>")


def test_str(sf: SheetFrame):
    assert str(sf).endswith("!$C$2:$D$6>")


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns == ["a", "b"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == []


def test_contains(sf: SheetFrame):
    assert "a" in sf
    assert "x" not in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ("a", 3),
        ("b", 4),
        (["a", "b"], [3, 4]),
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
        ("a", 0, "$C$3"),
        ("a", -1, "$C$2"),
        ("b", -1, "$D$2"),
        ("a", None, "$C$3:$C$6"),
    ],
)
def test_range(sf: SheetFrame, column, offset, address):
    assert sf.range(column, offset).get_address() == address


def test_range_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="invalid offset"):
        sf.range("a", 1)  # type: ignore


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


def test_setitem(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, style=False, sheet=sheet)
    x = [10, 20, 30]
    sf["a"] = x
    np.testing.assert_array_equal(sf["a"], x)


def test_setitem_new_column(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    sf = SheetFrame(2, 2, data=df, style=False, sheet=sheet)
    x = [10, 20, 30]
    sf["c"] = x
    assert sf.columns == [None, "a", "b", "c"]
    np.testing.assert_array_equal(sf["c"], x)


# def test_address(sf: SheetFrame):
#     s = sf.get_address("a")
#     assert isinstance(s, Series)
#     assert s.to_list() == ["$D$3", "$D$4", "$D$5", "$D$6"]
#     assert s.name == "a"


# def test_address_formula(sf: SheetFrame):
#     s = sf.get_address("a", formula=True)
#     assert s.to_list() == ["=$D$3", "=$D$4", "=$D$5", "=$D$6"]


# @pytest.mark.parametrize("columns", [["a", "b"], ["b", "a"], None])
# def test_address_list_or_none(sf: SheetFrame, columns):
#     df = sf.get_address(columns)
#     assert isinstance(df, DataFrame)
#     assert df.shape == (4, 2)
#     assert df["a"].to_list() == ["$D$3", "$D$4", "$D$5", "$D$6"]
#     assert df["b"].to_list() == ["$E$3", "$E$4", "$E$5", "$E$6"]
#     assert df.index.to_list() == list(range(4))
#     assert df.index.name is None


# def test_groupby(sheet: Sheet):
#     df = DataFrame({"a": [1, 1, 1, 2, 2, 1, 1], "b": [1, 2, 3, 4, 5, 6, 7]})
#     sf = SheetFrame(2, 2, data=df, style=False, index=False, sheet=sheet)

#     g = groupby(sf, "a")
#     assert len(g) == 2
#     assert g[(1,)] == [(3, 5), (8, 9)]
#     assert g[(2,)] == [(6, 7)]

#     assert len(groupby(sf, ["a", "b"])) == 7

#     g = groupby(sf, "::b")
#     assert len(g) == 2
#     assert g[(1,)] == [(3, 5), (8, 9)]
#     assert g[(2,)] == [(6, 7)]

#     assert len(groupby(sf, ":b")) == 7


# def test_groupby_range(sheet: Sheet):
#     df = DataFrame({"a": [1, 1, 1, 2, 2, 1, 1], "b": [3, 3, 4, 4, 3, 3, 4]})
#     sf = SheetFrame(2, 2, data=df, style=False, index=False, sheet=sheet)

#     g = groupby(sf, "a")
#     assert sf.range("a", g[(1,)]).get_address() == "$B$3:$B$5,$B$8:$B$9"
#     assert sf.range("a", g[(2,)]).get_address() == "$B$6:$B$7"

#     g = groupby(sf, "b")
#     assert sf.range("b", g[(3,)]).get_address() == "$C$3:$C$4,$C$7:$C$8"
#     assert sf.range("b", g[(4,)]).get_address() == "$C$5:$C$6,$C$9"

#     g = groupby(sf, ["a", "b"])
#     assert sf.range("a", g[(1, 3)]).get_address() == "$B$3:$B$4,$B$8"
#     assert sf.range("a", g[(1, 4)]).get_address() == "$B$5,$B$9"
#     assert sf.range("a", g[(2, 3)]).get_address() == "$B$7"
#     assert sf.range("a", g[(2, 4)]).get_address() == "$B$6"
