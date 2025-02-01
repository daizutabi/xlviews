import pytest

from xlviews.dataframes.groupby import GroupBy
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.range.range_collection import RangeCollection
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.mark.parametrize(
    ("by", "n"),
    [(None, 1), ("x", 2), (["x", "y"], 4), (["x", "y", "z"], 20)],
)
def test_by(sf: SheetFrame, by, n):
    gr = GroupBy(sf, by)
    assert len(gr.group) == n


@pytest.fixture(scope="module")
def gr(sf: SheetFrame):
    return GroupBy(sf, ["x", "y"])


def test_group_key(gr: GroupBy):
    keys = list(gr.group.keys())
    assert keys == [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]


def test_ranges_len(gr: GroupBy):
    assert len(list(gr.ranges("a"))) == 4


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_first_ranges(gr: GroupBy, column, c):
    rs = [r.get_address() for r in gr.first_ranges(column)]
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
def test_ranges(gr: GroupBy, column, c, k: int, a):
    rc = list(gr.ranges(column))[k]
    assert isinstance(rc, RangeCollection)
    assert rc.get_address() == a.format(c=c)
