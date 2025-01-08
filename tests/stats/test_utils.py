import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


def test_wrap_wrap():
    from xlviews.stats import get_wrap

    assert get_wrap(wrap="wrap") == "wrap"
    assert get_wrap(wrap={"a": "wrap"}) == {"a": "wrap"}


@pytest.mark.parametrize(
    ("kwarg", "value"),
    [("na", "IFERROR({},NA())"), ("null", 'IFERROR({},"")')],
)
def test_wrap_true(kwarg, value):
    from xlviews.stats import get_wrap

    assert get_wrap(**{kwarg: True}) == value  # type: ignore


def test_wrap_list():
    from xlviews.stats import get_wrap

    x = get_wrap(na=["a", "b"], null="c")
    assert isinstance(x, dict)
    assert x["a"] == "IFERROR({},NA())"
    assert x["b"] == "IFERROR({},NA())"
    assert x["c"] == 'IFERROR({},"")'


def test_wrap_none():
    from xlviews.stats import get_wrap

    assert get_wrap() is None


def test_func_none():
    from xlviews.stats import get_func

    func = ["count", "max", "mean", "median", "min", "soa"]
    assert sorted(get_func(None)) == func


def test_func_str():
    from xlviews.stats import get_func

    assert get_func("count") == ["count"]


@pytest.mark.parametrize("func", [["count"], {"a": "count"}])
def test_func_else(func):
    from xlviews.stats import get_func

    assert get_func(func) == func


def test_has_header(sf_parent: SheetFrame):
    from xlviews.stats import has_header

    assert has_header(sf_parent)


def test_move_down(sheet: Sheet):
    from xlviews.stats import move_down

    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(sheet, 3, 3, data=df, style=False)
    assert sheet["D3:F3"].value == ["a", "b", "c"]
    assert sheet["D2:F2"].value == [None, None, None]
    assert move_down(sf, 3) == 3
    assert sheet["D6:F6"].value == ["a", "b", "c"]
    assert sheet["D5:F5"].value == [None, None, None]
    assert sf.row == 6


def test_move_down_header(sheet: Sheet):
    from xlviews.stats import move_down

    df = DataFrame([[1, 2, 3], [4, 5, 6]], columns=["a", "b", "c"])
    sf = SheetFrame(sheet, 3, 3, data=df, style=False)
    sheet["D2"].value = "x"
    assert sheet["D3:F3"].value == ["a", "b", "c"]
    assert sheet["D2:F2"].value == ["x", None, None]
    assert move_down(sf, 3) == 4
    assert sheet["D7:F7"].value == ["a", "b", "c"]
    assert sheet["D6:F6"].value == ["x", None, None]
    assert sf.row == 7


def test_column_ranges(sheet_module: Sheet):
    from xlviews.stats import get_column_ranges

    rngs = get_column_ranges(sheet_module, [[1, 5], [9, 12]], 2)
    assert rngs[0].get_address() == "$B$1:$B$5"
    assert rngs[1].get_address() == "$B$9:$B$12"


def test_column_ranges_offset(sheet_module: Sheet):
    from xlviews.stats import get_column_ranges

    rngs = get_column_ranges(sheet_module, [[1, 5], [9, 12]], 2, 10)
    assert rngs[0].get_address() == "$B$11:$B$15"
    assert rngs[1].get_address() == "$B$19:$B$22"
