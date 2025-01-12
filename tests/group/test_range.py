import pytest

from xlviews.group import GroupedRange
from xlviews.range import RangeCollection
from xlviews.sheetframe import SheetFrame


@pytest.mark.parametrize(
    ("by", "n"),
    [(None, 1), ("x", 2), (["x", "y"], 4), (["x", "y", "z"], 20)],
)
def test_by(sf: SheetFrame, by, n):
    gr = GroupedRange(sf, by)
    assert len(gr.grouped) == n


@pytest.fixture(scope="module")
def gr(sf: SheetFrame):
    return GroupedRange(sf, ["x", "y"])


def test_group_key(gr: GroupedRange):
    keys = list(gr.grouped.keys())
    assert keys == [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]


def test_iter_ranges_len(gr: GroupedRange):
    assert len(list(gr.iter_ranges("a"))) == 4


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_first_ranges(gr: GroupedRange, column, c):
    rs = [r.get_address() for r in gr.iter_first_ranges(column)]
    assert rs == [f"${c}$4", f"${c}$8", f"${c}$12", f"${c}$16"]


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
@pytest.mark.parametrize(
    ("k", "a"),
    [
        (0, "${c}$4:${c}$7,${c}$20:${c}$23"),
        (1, "${c}$8:${c}$11"),
        (2, "${c}$12:${c}$15"),
        (3, "${c}$16:${c}$19"),
    ],
)
def test_iter_row_ranges_range(gr: GroupedRange, column, c, k: int, a):
    rc = list(gr.iter_ranges(column))[k]
    assert isinstance(rc, RangeCollection)
    assert rc.get_address() == a.format(c=c)
