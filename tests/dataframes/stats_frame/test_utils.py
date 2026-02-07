from __future__ import annotations

from typing import TYPE_CHECKING

import pytest
from pandas import DataFrame

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.stats_frame import get_func, get_length, has_header, move_down
from xlviews.testing import is_app_available

if TYPE_CHECKING:
    from xlwings import Sheet


def test_func_none():
    func = ["count", "max", "mean", "median", "min", "soa"]
    assert sorted(get_func(None)) == func


def test_func_str():
    assert get_func("count") == ["count"]


@pytest.mark.parametrize("func", [["count"]])
def test_func_else(func: list[str]):
    assert get_func(func) == func


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
@pytest.mark.parametrize(
    ("funcs", "n"),
    [(["mean"], 4), (["min", "max", "median"], 12)],
)
def test_length(sf_parent: SheetFrame, funcs: list[str], n: int):
    assert get_length(sf_parent, ["x", "y"], funcs) == n


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_length_none_list(sf_parent: SheetFrame):
    assert get_length(sf_parent, [], ["min", "max"]) == 2


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_has_header(sf_parent: SheetFrame):
    assert has_header(sf_parent)


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_move_down(sheet: Sheet):
    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(3, 3, data=df, sheet=sheet)
    assert sheet["D3:F3"].value == ["a", "b", "c"]
    assert sheet["D2:F2"].value == [None, None, None]
    assert move_down(sf, 3) == 3
    assert sheet["D6:F6"].value == ["a", "b", "c"]
    assert sheet["D5:F5"].value == [None, None, None]
    assert sf.row == 6


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_move_down_header(sheet: Sheet):
    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(3, 3, data=df, sheet=sheet)
    sheet["D2"].value = "x"
    assert sheet["D3:F3"].value == ["a", "b", "c"]
    assert sheet["D2:F2"].value == ["x", None, None]
    assert move_down(sf, 3) == 4
    assert sheet["D7:F7"].value == ["a", "b", "c"]
    assert sheet["D6:F6"].value == ["x", None, None]
    assert sf.row == 7
