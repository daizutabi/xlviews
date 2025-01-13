import string

import numpy as np
import pytest
from pandas import DataFrame, MultiIndex
from xlwings import Sheet

from xlviews.grouper import Grouper
from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def sf(sheet_module: Sheet):
    a = ["a"] * 8 + ["b"] * 8
    b = (["c"] * 4 + ["d"] * 4) * 2
    c = np.repeat(range(1, 9), 2)
    d = ["x", "y"] * 8
    df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
    df.columns = MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
    return SheetFrame(2, 2, data=df, index=True, style=False, sheet=sheet_module)


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
    gr = Grouper(sf, by)
    assert len(gr) == n


@pytest.fixture(scope="module")
def gr(sf: SheetFrame):
    return Grouper(sf, ["s", "t"])


def test_keys(gr: Grouper):
    keys = [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]
    assert list(gr.keys()) == keys


def test_values(gr: Grouper):
    values = [[(3, 6)], [(7, 10)], [(11, 14)], [(15, 18)]]
    assert list(gr.values()) == values


@pytest.mark.parametrize(
    ("column", "value"),
    [
        (("a", "c"), [2, 3, 4, 5]),
        (("a", "d"), [6, 7, 8, 9]),
        (("b", "c"), [10, 11, 12, 13]),
        (("b", "d"), [14, 15, 16, 17]),
    ],
)
def test_ranges(gr: Grouper, column: tuple, value):
    for rng, i in zip(gr.ranges(column), value, strict=True):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


def test_ranges_none(gr: Grouper):
    values = [[2, 3, 4, 5], [6, 7, 8, 9], [10, 11, 12, 13], [14, 15, 16, 17]]
    for it, value in zip(gr.ranges(), values, strict=True):
        for rng, i in zip(it, value, strict=True):
            c = string.ascii_uppercase[i]
            assert rng.get_address() == f"${c}$6:${c}$11"


def test_ranges_kwargs(gr: Grouper):
    values = [[2, 4], [6, 8], [10, 12], [14, 16]]
    for it, value in zip(gr.ranges(i="x"), values, strict=True):
        for rng, i in zip(it, value, strict=True):
            c = string.ascii_uppercase[i]
            assert rng.get_address() == f"${c}$6:${c}$11"
