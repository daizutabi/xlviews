import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available


def test_func_none():
    from xlviews.dataframes.stats_frame import get_func

    func = ["count", "max", "mean", "median", "min", "soa"]
    assert sorted(get_func(None)) == func


def test_func_str():
    from xlviews.dataframes.stats_frame import get_func

    assert get_func("count") == ["count"]


@pytest.mark.parametrize("func", [["count"], {"a": "count"}])
def test_func_else(func):
    from xlviews.dataframes.stats_frame import get_func

    assert get_func(func) == func


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
@pytest.mark.parametrize(
    ("funcs", "n"),
    [(["mean"], 4), (["min", "max", "median"], 12)],
)
def test_length(sf_parent: SheetFrame, funcs, n):
    from xlviews.dataframes.stats_frame import get_length

    assert get_length(sf_parent, ["x", "y"], funcs) == n


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_length_none_list(sf_parent: SheetFrame):
    from xlviews.dataframes.stats_frame import get_length

    assert get_length(sf_parent, [], ["min", "max"]) == 2


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_has_header(sf_parent: SheetFrame):
    from xlviews.dataframes.stats_frame import has_header

    assert has_header(sf_parent)


@pytest.mark.skipif(not is_app_available(), reason="Excel not installed")
def test_move_down(sheet: Sheet):
    from xlviews.dataframes.stats_frame import move_down

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
    from xlviews.dataframes.stats_frame import move_down

    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(3, 3, data=df, sheet=sheet)
    sheet["D2"].value = "x"
    assert sheet["D3:F3"].value == ["a", "b", "c"]
    assert sheet["D2:F2"].value == ["x", None, None]
    assert move_down(sf, 3) == 4
    assert sheet["D7:F7"].value == ["a", "b", "c"]
    assert sheet["D6:F6"].value == ["x", None, None]
    assert sf.row == 7
