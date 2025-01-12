import numpy as np
import pytest
from pandas import DataFrame, MultiIndex
from xlwings import Sheet

from xlviews.group import GroupedRange
from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def sf(sheet_module: Sheet):
    a = ["a"] * 8 + ["b"] * 8
    b = (["c"] * 4 + ["d"] * 4) * 2
    c = np.repeat(range(1, 9), 2)
    d = ["x", "y"] * 8
    df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
    df.columns = MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
    return SheetFrame(sheet_module, 2, 2, data=df, index=True, style=False)


@pytest.mark.parametrize(
    ("by", "n"),
    [
        ("s", 2),
        ("t", 2),
        ("r", 8),
        ("i", 2),
        (["s", "t"], 4),
        (["s", "t", "i"], 8),
        (["t", "i"], 4),
        (["s", "t", "r", "i"], 16),
    ],
)
def test_len(sf: SheetFrame, by, n: int):
    gr = GroupedRange(sf, by)
    assert len(gr) == n


@pytest.fixture(scope="module")
def gr(sf: SheetFrame):
    return GroupedRange(sf, ["s", "t"])


def test_keys(gr: GroupedRange):
    keys = [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]
    assert list(gr.keys()) == keys


def test_values(gr: GroupedRange):
    values = [[(3, 6)], [(7, 10)], [(11, 14)], [(15, 18)]]
    assert list(gr.values()) == values


# @pytest.mark.parametrize(
#     ("key", "value"),
#     [
#         (("c", 100), "$E$3:$E$4,$E$8:$E$9"),
#         (("c", 200), "$E$5:$E$7,$E$10:$E$12"),
#     ],
# )
# def test_range(gr: GroupedRange, key, value):
#     assert gr.range("x", key).get_address() == value


# @pytest.mark.parametrize(
#     ("key", "value"),
#     [
#         (("c", 100), "$B$3"),
#         (("c", 200), "$B$5"),
#     ],
# )
# def test_first_range(gr: GroupedRange, key, value):
#     assert gr.first_range("a", key).get_address() == value
