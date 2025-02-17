import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.groupby import GroupBy, to_dict
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize(
    ("keys", "values", "expected"),
    [
        ([1, 2, 1], ["a", "b", "c"], {1: ["a", "c"], 2: ["b"]}),
        (["x", "y", "x"], [10, 20, 30], {"x": [10, 30], "y": [20]}),
        ([True, False, True], [1.1, 2.2, 3.3], {True: [1.1, 3.3], False: [2.2]}),
        ([], [], {}),
        ([None, None], ["a", "b"], {None: ["a", "b"]}),
    ],
)
def test_to_dict(keys, values, expected):
    assert to_dict(keys, values) == expected


@pytest.mark.parametrize("func", [lambda x: x, np.array, Series])
def test_create_group_index_series(func):
    from xlviews.dataframes.groupby import create_group_index

    values = [1, 1, 2, 2, 2, 3, 3, 1, 1, 2, 2, 3, 3]
    index = create_group_index(func(values))
    assert index[(1,)] == [(0, 1), (7, 8)]
    assert index[(2,)] == [(2, 4), (9, 10)]
    assert index[(3,)] == [(5, 6), (11, 12)]


@pytest.mark.parametrize("func", [lambda x: x, DataFrame])
def test_create_group_index_dataframe(func):
    from xlviews.dataframes.groupby import create_group_index

    values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
    index = create_group_index(func(values))
    assert index[(1, 2)] == [(0, 1), (4, 4)]
    assert index[(3, 4)] == [(2, 3), (5, 6)]


@pytest.fixture(scope="module")
def sf(sheet_module: Sheet):
    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    return SheetFrame(2, 2, df, sheet_module)


@pytest.mark.parametrize(
    ("by", "n"),
    [
        ("a", 1),
        ("b", 2),
        ("c", 2),
        (["a", "b"], 2),
        (["a", "c"], 2),
        (["b", "c"], 4),
        (["a", "b", "c"], 4),
    ],
)
def test_len(sf: SheetFrame, by, n: int):
    gr = GroupBy(sf, by)
    assert len(gr) == n


@pytest.fixture(scope="module")
def gp(sf: SheetFrame):
    return GroupBy(sf, ["a", "c"])


def test_keys(gp: GroupBy):
    keys = [("c", 100), ("c", 200)]
    assert list(gp.keys()) == keys


def test_values(gp: GroupBy):
    values = [[(3, 4), (8, 9)], [(5, 7), (10, 12)]]
    assert list(gp.values()) == values


def test_items(gp: GroupBy):
    assert next(gp.items()) == (("c", 100), [(3, 4), (8, 9)])


def test_iter(gp: GroupBy):
    assert next(iter(gp)) == ("c", 100)


@pytest.mark.parametrize(
    ("key", "value"),
    [
        (("c", 100), [(3, 4), (8, 9)]),
        (("c", 200), [(5, 7), (10, 12)]),
    ],
)
def test_getitem(gp: GroupBy, key, value):
    assert gp[key] == value


@pytest.fixture
def sf2(sheet: Sheet):
    df = DataFrame(
        {
            "x": ["a"] * 8 + ["b"] * 8 + ["a"] * 4,
            "y": (["c"] * 4 + ["d"] * 4) * 2 + ["c"] * 4,
            "z": range(1, 21),
            "a": range(1, 21),
        },
    )
    df = df.set_index(["x", "y", "z"])
    return SheetFrame(3, 3, data=df, sheet=sheet)


@pytest.mark.parametrize(
    ("by", "n"),
    [(None, 1), ("x", 2), (["x", "y"], 4), (["x", "y", "z"], 20)],
)
def test_by(sf2: SheetFrame, by, n):
    gp = GroupBy(sf2, by)
    assert len(gp.group) == n


def test_group_key(sf2: SheetFrame):
    gp = GroupBy(sf2, ["x", "y"])
    keys = list(gp.group.keys())
    assert keys == [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]
