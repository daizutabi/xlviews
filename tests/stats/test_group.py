import pytest
from xlwings import Range

from xlviews.frame import SheetFrame
from xlviews.stats import GroupedRange


@pytest.mark.parametrize(
    ("by", "n"),
    [(None, 1), ("x", 2), (["x", "y"], 4), (["x", "y", "z"], 20)],
)
def test_by(sf_parent: SheetFrame, by, n):
    gr = GroupedRange(sf_parent, by)
    assert len(gr.grouped) == n


@pytest.fixture(scope="module")
def gr(sf_parent: SheetFrame):
    return GroupedRange(sf_parent, ["x", "y"])


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_row_ranges_str(gr: GroupedRange, column, c):
    rs = list(gr.iter_row_ranges(column))
    assert rs == [f"${c}$3", f"${c}$7", f"${c}$11", f"${c}$15"]


def test_iter_row_ranges_none(gr: GroupedRange):
    rs = list(gr.iter_row_ranges("z"))
    assert rs == ["", "", "", ""]


def test_group_key(gr: GroupedRange):
    keys = list(gr.grouped.keys())
    assert keys == [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]


def test_iter_row_ranges_len(gr: GroupedRange):
    rs = list(gr.iter_row_ranges("a"))
    assert len(rs) == 4


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
@pytest.mark.parametrize(
    ("k", "r"),
    [(0, [3, 6]), (1, [7, 10]), (2, [11, 14]), (3, [15, 18])],
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
        assert x.get_address() == f"${c}$19:${c}$22"
    else:
        assert len(rs[k]) == 1


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_formulas_list_index(gr: GroupedRange, column, c):
    fs = list(gr.iter_formulas(column, ["min", "max"]))
    a = [f"=${c}${x}" for x in ["3", "3", "7", "7", "11", "11", "15", "15"]]
    assert fs == a


def test_iter_formulas_list_index_none(gr: GroupedRange):
    fs = list(gr.iter_formulas("z", ["min", "max"]))
    assert fs == [""] * 8


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
def test_iter_formulas_list_values(gr: GroupedRange, column, c):
    fs = list(gr.iter_formulas(column, ["min", "max"], wrap="__{}__"))
    assert fs[0] == f"=__AGGREGATE(5,7,${c}$3:${c}$6,${c}$19:${c}$22)__"
    assert fs[1] == f"=__AGGREGATE(4,7,${c}$3:${c}$6,${c}$19:${c}$22)__"
    assert fs[2] == f"=__AGGREGATE(5,7,${c}$7:${c}$10)__"
    assert fs[3] == f"=__AGGREGATE(4,7,${c}$7:${c}$10)__"
    assert fs[4] == f"=__AGGREGATE(5,7,${c}$11:${c}$14)__"
    assert fs[5] == f"=__AGGREGATE(4,7,${c}$11:${c}$14)__"
    assert fs[6] == f"=__AGGREGATE(5,7,${c}$15:${c}$18)__"
    assert fs[7] == f"=__AGGREGATE(4,7,${c}$15:${c}$18)__"


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_formulas_dict_index(gr: GroupedRange, column, c):
    fs = list(gr.iter_formulas(column, {}))
    a = [f"=${c}${x}" for x in ["3", "7", "11", "15"]]
    assert fs == a


def test_iter_formulas_dict_index_none(gr: GroupedRange):
    fs = list(gr.iter_formulas("z", {}))
    assert fs == [""] * 4


@pytest.mark.parametrize(
    ("column", "c", "k"),
    [("a", "F", 1), ("b", "G", 12), ("c", "H", 9)],
)
def test_iter_formulas_dict_values(gr: GroupedRange, column, c, k):
    funcs = {"a": "mean", "b": "median", "c": "sum"}
    fs = list(gr.iter_formulas(column, funcs, wrap="__{}__"))
    assert fs[0] == f"=__AGGREGATE({k},7,${c}$3:${c}$6,${c}$19:${c}$22)__"
    assert fs[1] == f"=__AGGREGATE({k},7,${c}$7:${c}$10)__"
    assert fs[2] == f"=__AGGREGATE({k},7,${c}$11:${c}$14)__"
    assert fs[3] == f"=__AGGREGATE({k},7,${c}$15:${c}$18)__"


def test_get_index(gr: GroupedRange):
    index = gr.get_index(["a", "b"])
    assert index == ["a", "b", "a", "b", "a", "b", "a", "b"]
