import pytest
from xlwings import Range

from xlviews.group import GroupedRange
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


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_row_ranges_str(gr: GroupedRange, column, c):
    rs = list(gr.iter_row_ranges(column))
    assert rs == [f"${c}$4", f"${c}$8", f"${c}$12", f"${c}$16"]


def test_iter_row_ranges_none(gr: GroupedRange):
    rs = list(gr.iter_row_ranges("z"))
    assert rs == ["", "", "", ""]


def test_iter_row_ranges_len(gr: GroupedRange):
    rs = list(gr.iter_row_ranges("a"))
    assert len(rs) == 4


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
@pytest.mark.parametrize(
    ("k", "r"),
    [(0, [4, 7]), (1, [8, 11]), (2, [12, 15]), (3, [16, 19])],
)
def test_iter_row_ranges_range(gr: GroupedRange, column, c, k, r):
    rs = list(gr.iter_row_ranges(column))
    x = rs[k][0]
    assert isinstance(x, Range)
    assert x.get_address() == f"${c}${r[0]}:${c}${r[1]}"
    if k == 0:
        assert len(rs[k]) == 2
        x = rs[k][1]
        assert isinstance(x, Range)
        assert x.get_address() == f"${c}$20:${c}$23"
    else:
        assert len(rs[k]) == 1
