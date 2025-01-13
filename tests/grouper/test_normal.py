import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.grouper import Grouper
from xlviews.sheetframe import SheetFrame


@pytest.fixture(scope="module")
def sf(sheet_module: Sheet):
    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    return SheetFrame(2, 2, data=df, index=True, style=False, sheet=sheet_module)


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
    gr = Grouper(sf, by)
    assert len(gr) == n


@pytest.fixture(scope="module")
def gr(sf: SheetFrame):
    return Grouper(sf, ["a", "c"])


def test_keys(gr: Grouper):
    keys = [("c", 100), ("c", 200)]
    assert list(gr.keys()) == keys


def test_values(gr: Grouper):
    values = [[(3, 4), (8, 9)], [(5, 7), (10, 12)]]
    assert list(gr.values()) == values


def test_items(gr: Grouper):
    assert next(gr.items()) == (("c", 100), [(3, 4), (8, 9)])


def test_iter(gr: Grouper):
    assert next(iter(gr)) == ("c", 100)


@pytest.mark.parametrize(
    ("key", "value"),
    [
        (("c", 100), [(3, 4), (8, 9)]),
        (("c", 200), [(5, 7), (10, 12)]),
    ],
)
def test_getitem(gr: Grouper, key, value):
    assert gr[key] == value


@pytest.mark.parametrize(
    ("key", "value"),
    [
        (("c", 100), "$E$3:$E$4,$E$8:$E$9"),
        (("c", 200), "$E$5:$E$7,$E$10:$E$12"),
    ],
)
def test_range(gr: Grouper, key, value):
    assert gr.range("x", key).get_address() == value


@pytest.mark.parametrize(
    ("key", "value"),
    [
        (("c", 100), "$B$3"),
        (("c", 200), "$B$5"),
    ],
)
def test_first_range(gr: Grouper, key, value):
    assert gr.first_range("a", key).get_address() == value
