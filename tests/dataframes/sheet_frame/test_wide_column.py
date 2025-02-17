import pytest
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import WideColumn

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return WideColumn(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 3
    assert sf.column == 29
    assert sf.sheet.name == sheet_module.name
    assert sf.index.nlevels == 2
    assert sf.columns.nlevels == 1


def test_len(sf: SheetFrame):
    assert len(sf) == 5


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["x", "y"]


def test_contains(sf: SheetFrame):
    assert "x" not in sf
    assert "a" in sf
    assert "u" in sf


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b", "u", "v"]


# @pytest.mark.parametrize(
#     ("column", "index"),
#     [
#         ("a", 31),
#         ("b", 32),
#         (["y", "b"], [30, 32]),
#         ("u", (33, 35)),
#         ("v", (36, 39)),
#         (("v", 0), 36),
#         (("v", 3), 39),
#         (["x", "a", "u", ("v", 0)], [29, 31, (33, 35), 36]),
#     ],
# )
# def test_index(sf: SheetFrame, column, index):
#     assert sf.index_past(column) == index


# @pytest.mark.parametrize("column", ["z", ("u", -1)])
# def test_index_error(sf: SheetFrame, column):
#     with pytest.raises(ValueError, match=".* is not in list"):
#         sf.index_past(column)


# @pytest.mark.parametrize(
#     ("column", "offset", "address"),
#     [
#         ("x", -1, "$AC$3"),
#         ("y", 0, "$AD$4"),
#         ("a", None, "$AE$4:$AE$8"),
#         (("u", 0), -1, "$AG$2:$AG$3"),
#         (("u", 2), 0, "$AI$4"),
#         (("v", 0), None, "$AJ$4:$AJ$8"),
#         ("u", -1, "$AG$2:$AI$3"),
#         ("u", 0, "$AG$4:$AI$4"),
#         ("v", None, "$AJ$4:$AM$8"),
#     ],
# )
# def test_range_column(sf: SheetFrame, column, offset, address):
#     assert sf.range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("by", "v1", "v2"),
    [
        ("x", [(4, 5), (8, 8)], [(6, 7)]),
        ("y", [(4, 4), (6, 6), (8, 8)], [(5, 5), (7, 7)]),
    ],
)
def test_groupby(sf: SheetFrame, by, v1, v2):
    g = groupby(sf, by)
    assert len(g) == 2
    keys = list(g.keys())
    assert g[keys[0]] == v1
    assert g[keys[1]] == v2


def test_groupby_list(sf: SheetFrame):
    g = groupby(sf, ["x", "y"])
    assert len(g) == 4
    assert g[("i", "k")] == [(4, 4), (8, 8)]
    assert g[("i", "l")] == [(5, 5)]
    assert g[("j", "k")] == [(6, 6)]
    assert g[("j", "l")] == [(7, 7)]
