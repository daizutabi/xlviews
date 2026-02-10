from __future__ import annotations

from typing import TYPE_CHECKING, Literal

import pytest
from pandas import DataFrame, Series

from xlviews.core.range import Range
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import NoIndex

if TYPE_CHECKING:
    from xlwings import App, Sheet

    from xlviews.dataframes.sheet_frame import SheetFrame

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return NoIndex(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, fc: FrameContainer):
    assert sf.row == fc.row
    assert sf.column == fc.column
    assert sf.index.nlevels == 1
    assert sf.columns.nlevels == 1


def test_repr(sf: SheetFrame):
    assert repr(sf).endswith("!$B$2:$D$6>")


def test_str(sf: SheetFrame):
    assert str(sf).endswith("!$B$2:$D$6>")


def test_len(sf: SheetFrame):
    assert len(sf) == 4


@pytest.mark.parametrize(("x", "b"), [("a", True), ("x", False), (0, False)])
def test_contains(sf: SheetFrame, x: str, b: bool):
    assert (x in sf) is b


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


def test_value(sf: SheetFrame, df: DataFrame):
    df_sf = sf.value
    assert df_sf.equals(df.astype(float))
    assert df_sf.index.equals(df.index)
    assert df_sf.columns.equals(df.columns)


@pytest.mark.parametrize(("column", "loc"), [("a", 3), ("b", 4)])
def test_loc(sf: SheetFrame, column: str, loc: int):
    assert sf.get_loc(column) == loc


@pytest.mark.parametrize(
    ("columns", "indexer"),
    [(["a"], [3]), (["b"], [4]), (None, [3, 4])],
)
def test_get_indexer(sf: SheetFrame, columns: list[str] | None, indexer: list[int]):
    assert sf.get_indexer(columns).tolist() == indexer


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("a", 0, "$C$3"),
        ("a", -1, "$C$2"),
        ("b", -1, "$D$2"),
        ("a", None, "$C$3:$C$6"),
    ],
)
def test_get_range(
    sf: SheetFrame,
    column: str,
    offset: Literal[0, -1] | None,
    address: str,
):
    assert sf.get_range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("axis", "v0", "v1"),
    [(0, [1, 2, 3, 4], [5, 6, 7, 8]), (1, [1, 5], [2, 6])],
)
def test_iter_ranges(sf: SheetFrame, axis: Literal[0, 1], v0: list[int], v1: list[int]):
    values = list(sf.iter_ranges(axis))
    assert values[0].value == v0
    assert values[1].value == v1


@pytest.mark.parametrize("columns", [["a", "b"], None])
def test_get_address_none(sf: SheetFrame, columns: list[str] | None):
    df = sf.get_address(columns)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.to_list() == [0, 1, 2, 3]
    assert df.loc[0, "a"] == "$C$3"
    assert df.loc[0, "b"] == "$D$3"
    assert df.loc[1, "a"] == "$C$4"
    assert df.loc[1, "b"] == "$D$4"
    assert df.loc[2, "a"] == "$C$5"
    assert df.loc[2, "b"] == "$D$5"
    assert df.loc[3, "a"] == "$C$6"
    assert df.loc[3, "a"] == "$C$6"


def test_get_address_str(sf: SheetFrame):
    s = sf.get_address("a", formula=True, row_absolute=False)
    assert s.name == "a"
    assert s.index.to_list() == [0, 1, 2, 3]
    assert s.to_list() == ["=$C3", "=$C4", "=$C5", "=$C6"]


def test_agg_str(sf: SheetFrame):
    s = sf.agg("sum", row_absolute=False, column_absolute=False)
    assert s.name == "sum"
    assert s.index.to_list() == ["a", "b"]
    assert s.to_list() == ["AGGREGATE(9,7,C3:C6)", "AGGREGATE(9,7,D3:D6)"]


def test_agg_dict(sf: SheetFrame):
    s = sf.agg({"a": "min", "b": "max"}, row_absolute=False, column_absolute=False)
    assert s.name is None
    assert s.index.to_list() == ["a", "b"]
    assert s.to_list() == ["AGGREGATE(5,7,C3:C6)", "AGGREGATE(4,7,D3:D6)"]


def test_agg_none(sf: SheetFrame):
    s = sf.agg()
    assert s.name is None
    assert s.index.to_list() == ["a", "b"]
    assert s.to_list() == ["$C$3:$C$6", "$D$3:$D$6"]


def test_agg_range(sf: SheetFrame):
    func = Range(100, 200, sf.sheet)
    s = sf.agg(func)
    assert s.name is None
    assert s.index.to_list() == ["a", "b"]
    assert 'IF(GR100="soa"' in s["a"]


def test_agg_range_error_sheet(sf: SheetFrame, sheet: Sheet):
    func = Range(100, 200, sheet)
    with pytest.raises(ValueError, match="Range is from a different sheet"):
        sf.agg(func)


def test_agg_range_error_book(sf: SheetFrame, app: App):
    book = app.books.add()

    func = Range(100, 200, book.sheets[0])
    with pytest.raises(ValueError, match="Range is from a different book"):
        sf.agg(func)

    book.close()


def test_agg_list(sf: SheetFrame):
    df = sf.agg(["median", "mean"])
    assert df.index.to_list() == ["median", "mean"]
    assert df.columns.to_list() == ["a", "b"]
    assert df.loc["median", "a"] == "AGGREGATE(12,7,$C$3:$C$6)"
    assert df.loc["mean", "b"] == "AGGREGATE(1,7,$D$3:$D$6)"


def test_agg_first(sf: SheetFrame):
    s = sf.agg("first")
    assert s.name == "first"
    assert s.to_list() == ["$C$3", "$D$3"]


def test_agg_columns_str(sf: SheetFrame):
    s = sf.agg(None, "a")
    assert s.name is None
    assert s.index.to_list() == ["a"]
    assert s.to_list() == ["$C$3:$C$6"]


def test_melt_none(sf: SheetFrame):
    s = sf.melt()
    assert isinstance(s, Series)
    assert s.name is None
    assert s.index.to_list() == ["a", "b"]
    assert s.to_list() == ["$C$3:$C$6", "$D$3:$D$6"]


def test_melt_str(sf: SheetFrame):
    s = sf.melt("std")
    assert isinstance(s, Series)
    assert s.name == "std"
    assert s.index.to_list() == ["a", "b"]
    assert s.to_list() == ["AGGREGATE(8,7,$C$3:$C$6)", "AGGREGATE(8,7,$D$3:$D$6)"]
