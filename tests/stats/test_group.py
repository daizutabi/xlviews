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


def test_group_key(gr: GroupedRange):
    keys = list(gr.grouped.keys())
    assert keys == [("a", "c"), ("a", "d"), ("b", "c"), ("b", "d")]


@pytest.mark.parametrize(
    ("funcs", "n"),
    [(["mean"], 4), (["min", "max", "median"], 12), ({"a": "count"}, 4)],
)
def test_length(gr: GroupedRange, funcs, n):
    assert gr.get_length(funcs) == n


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
@pytest.mark.parametrize("o", [0, 10])
def test_iter_row_ranges_str(gr: GroupedRange, column, c, o):
    rs = list(gr.iter_row_ranges(column, o))
    assert rs == [f"${c}${4+o}", f"${c}${8+o}", f"${c}${12+o}", f"${c}${16+o}"]


@pytest.mark.parametrize("o", [0, 10])
def test_iter_row_ranges_none(gr: GroupedRange, o):
    rs = list(gr.iter_row_ranges("z", o))
    assert rs == ["", "", "", ""]


def test_iter_row_ranges_len(gr: GroupedRange):
    rs = list(gr.iter_row_ranges("a"))
    assert len(rs) == 4


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
@pytest.mark.parametrize(
    ("k", "r"),
    [(0, [4, 7]), (1, [8, 11]), (2, [12, 15]), (3, [16, 19])],
)
@pytest.mark.parametrize("o", [0, 5])
def test_iter_row_ranges_range(gr: GroupedRange, column, c, k, r, o):
    rs = list(gr.iter_row_ranges(column, o))
    x = rs[k][0]
    assert isinstance(x, Range)
    assert x.get_address() == f"${c}${r[0]+o}:${c}${r[1]+o}"
    if k == 0:
        assert len(rs[k]) == 2
        x = rs[k][1]
        assert isinstance(x, Range)
        assert x.get_address() == f"${c}${20+o}:${c}${23+o}"
    else:
        assert len(rs[k]) == 1


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
@pytest.mark.parametrize("o", [0, 20])
def test_iter_formulas_list_index(gr: GroupedRange, column, c, o):
    fs = list(gr.iter_formulas(column, ["min", "max"], offset=o))
    a = [f"=${c}${x+o}" for x in [4, 4, 8, 8, 12, 12, 16, 16]]
    assert fs == a


@pytest.mark.parametrize("o", [0, 20])
def test_iter_formulas_list_index_none(gr: GroupedRange, o):
    fs = list(gr.iter_formulas("z", ["min", "max"], offset=o))
    assert fs == [""] * 8


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
@pytest.mark.parametrize("o", [0, 6])
def test_iter_formulas_list_columns(gr: GroupedRange, column, c, o):
    fs = list(gr.iter_formulas(column, ["min", "max"], wrap="__{}__", offset=o))
    assert len(fs) == 8
    assert fs[0] == f"=__AGGREGATE(5,7,${c}${4+o}:${c}${7+o},${c}${20+o}:${c}${23+o})__"
    assert fs[1] == f"=__AGGREGATE(4,7,${c}${4+o}:${c}${7+o},${c}${20+o}:${c}${23+o})__"
    assert fs[2] == f"=__AGGREGATE(5,7,${c}${8+o}:${c}${11+o})__"
    assert fs[3] == f"=__AGGREGATE(4,7,${c}${8+o}:${c}${11+o})__"
    assert fs[4] == f"=__AGGREGATE(5,7,${c}${12+o}:${c}${15+o})__"
    assert fs[5] == f"=__AGGREGATE(4,7,${c}${12+o}:${c}${15+o})__"
    assert fs[6] == f"=__AGGREGATE(5,7,${c}${16+o}:${c}${19+o})__"
    assert fs[7] == f"=__AGGREGATE(4,7,${c}${16+o}:${c}${19+o})__"


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
@pytest.mark.parametrize("o", [0, 16])
def test_iter_formulas_dict_index(gr: GroupedRange, column, c, o):
    fs = list(gr.iter_formulas(column, {}, offset=o))
    a = [f"=${c}${x+o}" for x in [4, 8, 12, 16]]
    assert fs == a


@pytest.mark.parametrize("o", [0, 16])
def test_iter_formulas_dict_index_none(gr: GroupedRange, o):
    fs = list(gr.iter_formulas("z", {}, offset=o))
    assert fs == [""] * 4


@pytest.mark.parametrize(
    ("column", "c", "k"),
    [("a", "F", 1), ("b", "G", 12), ("c", "H", 9)],
)
@pytest.mark.parametrize("o", [0, 16])
def test_iter_formulas_dict_columns(gr: GroupedRange, column, c, k, o):
    funcs = {"a": "mean", "b": "median", "c": "sum"}
    fs = list(gr.iter_formulas(column, funcs, wrap="__{}__", offset=o))
    assert len(fs) == 4
    x = f"=__AGGREGATE({k},7,${c}${4+o}:${c}${7+o},${c}${20+o}:${c}${23+o})__"
    assert fs[0] == x
    assert fs[1] == f"=__AGGREGATE({k},7,${c}${8+o}:${c}${11+o})__"
    assert fs[2] == f"=__AGGREGATE({k},7,${c}${12+o}:${c}${15+o})__"
    assert fs[3] == f"=__AGGREGATE({k},7,${c}${16+o}:${c}${19+o})__"


def test_get_index(gr: GroupedRange):
    index = gr.get_index(["a", "b"])
    assert index == ["a", "b", "a", "b", "a", "b", "a", "b"]


def test_get_columns_list(gr: GroupedRange):
    columns = gr.get_columns([])
    assert columns == ["func", "x", "y", "z", "a", "b", "c"]


def test_get_columns_dict(gr: GroupedRange):
    columns = gr.get_columns({"a": "mean", "b": "median"})
    assert columns == ["x", "y", "z", "a", "b", "c"]


@pytest.mark.parametrize(
    ("funcs", "shape"),
    [(["mean"], (4, 7)), (["min", "max", "median"], (12, 7)), ({"a": "count"}, (4, 6))],
)
def test_get_values_shape(gr: GroupedRange, funcs, shape):
    assert gr.get_values(funcs).shape == shape


@pytest.mark.parametrize(
    ("funcs", "shape"),
    [(["mean"], (4, 3)), (["min", "max"], (8, 3)), ({"a": "std"}, (4, 3))],
)
def test_get_frame_shape(gr: GroupedRange, funcs, shape):
    assert gr.get_frame(funcs).shape == shape


def test_get_frame_index_list(gr: GroupedRange):
    df = gr.get_frame(["mean"])
    assert df.index.names == ["func", "x", "y", "z"]


def test_get_frame_index_dict(gr: GroupedRange):
    df = gr.get_frame({})
    assert df.index.names == ["x", "y", "z"]


@pytest.mark.parametrize("funcs", [[], {}])
def test_get_frame_columns(gr: GroupedRange, funcs):
    df = gr.get_frame(funcs)
    assert df.columns.to_list() == ["a", "b", "c"]


def test_get_frame_wrap_str(gr: GroupedRange):
    df = gr.get_frame(["mean"], wrap="__{}__")
    values = df.to_numpy().flatten()
    assert all(x.startswith("=__AGGREGATE") for x in values)
    assert all(x.endswith(")__") for x in values)


def test_get_frame_wrap_dict(gr: GroupedRange):
    df = gr.get_frame(["mean"], wrap={"a": "A{}A", "b": "B{}B"})
    assert all(x.startswith("=AAGGREGATE") for x in df["a"])
    assert all(x.endswith(")A") for x in df["a"])
    assert all(x.startswith("=BAGGREGATE") for x in df["b"])
    assert all(x.endswith(")B") for x in df["b"])
    assert all(x.startswith("=AGGREGATE") for x in df["c"])
    assert all(x.endswith(")") for x in df["c"])


@pytest.mark.parametrize("o", [0, 16])
def test_get_frame_offset(gr: GroupedRange, o):
    df = gr.get_frame(["mean"], offset=o).reset_index()
    assert df["x"].iloc[0] == f"=$C${4+o}"
    assert df["y"].iloc[-1] == f"=$D${16+o}"
    assert df["a"].iloc[0] == f"=AGGREGATE(1,7,$F${4+o}:$F${7+o},$F${20+o}:$F${23+o})"
    assert df["c"].iloc[-1] == f"=AGGREGATE(1,7,$H${16+o}:$H${19+o})"
