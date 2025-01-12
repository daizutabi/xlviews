import pytest

from xlviews.sheetframe import SheetFrame
from xlviews.statsframe import StatsGrouper


@pytest.fixture(scope="module")
def gr(sf_parent: SheetFrame):
    return StatsGrouper(sf_parent, ["x", "y"])


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_formulas_list_index(gr: StatsGrouper, column, c):
    fs = list(gr.iter_formulas(column, ["min", "max"]))
    a = [f"=${c}${x}" for x in [4, 4, 8, 8, 12, 12, 16, 16]]
    assert fs == a


def test_iter_formulas_list_index_none(gr: StatsGrouper):
    fs = list(gr.iter_formulas("z", ["min", "max"]))
    assert fs == [""] * 8


@pytest.mark.parametrize(("column", "c"), [("a", "F"), ("b", "G"), ("c", "H")])
def test_iter_formulas_list_columns(gr: StatsGrouper, column, c):
    fs = list(gr.iter_formulas(column, ["min", "max"], wrap="__{}__"))
    assert len(fs) == 8
    assert fs[0] == f"=__AGGREGATE(5,7,${c}$4:${c}$7,${c}$20:${c}$23)__"
    assert fs[1] == f"=__AGGREGATE(4,7,${c}$4:${c}$7,${c}$20:${c}$23)__"
    assert fs[2] == f"=__AGGREGATE(5,7,${c}$8:${c}$11)__"
    assert fs[3] == f"=__AGGREGATE(4,7,${c}$8:${c}$11)__"
    assert fs[4] == f"=__AGGREGATE(5,7,${c}$12:${c}$15)__"
    assert fs[5] == f"=__AGGREGATE(4,7,${c}$12:${c}$15)__"
    assert fs[6] == f"=__AGGREGATE(5,7,${c}$16:${c}$19)__"
    assert fs[7] == f"=__AGGREGATE(4,7,${c}$16:${c}$19)__"


@pytest.mark.parametrize(("column", "c"), [("x", "C"), ("y", "D")])
def test_iter_formulas_dict_index(gr: StatsGrouper, column, c):
    fs = list(gr.iter_formulas(column, {}))
    a = [f"=${c}${x}" for x in [4, 8, 12, 16]]
    assert fs == a


def test_iter_formulas_dict_index_none(gr: StatsGrouper):
    fs = list(gr.iter_formulas("z", {}))
    assert fs == [""] * 4


@pytest.mark.parametrize(
    ("column", "c", "k"),
    [("a", "F", 1), ("b", "G", 12), ("c", "H", 9)],
)
def test_iter_formulas_dict_columns(gr: StatsGrouper, column, c, k):
    funcs = {"a": "mean", "b": "median", "c": "sum"}
    fs = list(gr.iter_formulas(column, funcs, wrap="__{}__"))
    assert len(fs) == 4
    x = f"=__AGGREGATE({k},7,${c}$4:${c}$7,${c}$20:${c}$23)__"
    assert fs[0] == x
    assert fs[1] == f"=__AGGREGATE({k},7,${c}$8:${c}$11)__"
    assert fs[2] == f"=__AGGREGATE({k},7,${c}$12:${c}$15)__"
    assert fs[3] == f"=__AGGREGATE({k},7,${c}$16:${c}$19)__"


def test_get_index(gr: StatsGrouper):
    index = gr.get_index(["a", "b"])
    assert index == ["a", "b", "a", "b", "a", "b", "a", "b"]


def test_get_columns_list(gr: StatsGrouper):
    columns = gr.get_columns([])
    assert columns == ["func", "x", "y", "z", "a", "b", "c"]


def test_get_columns_dict(gr: StatsGrouper):
    columns = gr.get_columns({"a": "mean", "b": "median"})
    assert columns == ["x", "y", "z", "a", "b", "c"]


@pytest.mark.parametrize(
    ("funcs", "shape"),
    [(["mean"], (4, 7)), (["min", "max", "median"], (12, 7)), ({"a": "count"}, (4, 6))],
)
def test_get_values_shape(gr: StatsGrouper, funcs, shape):
    assert gr.get_values(funcs).shape == shape


@pytest.mark.parametrize(
    ("funcs", "shape"),
    [(["mean"], (4, 3)), (["min", "max"], (8, 3)), ({"a": "std"}, (4, 3))],
)
def test_get_frame_shape(gr: StatsGrouper, funcs, shape):
    assert gr.get_frame(funcs).shape == shape


def test_get_frame_index_list(gr: StatsGrouper):
    df = gr.get_frame(["mean"])
    assert df.index.names == ["func", "x", "y", "z"]


def test_get_frame_index_dict(gr: StatsGrouper):
    df = gr.get_frame({})
    assert df.index.names == ["x", "y", "z"]


@pytest.mark.parametrize("funcs", [[], {}])
def test_get_frame_columns(gr: StatsGrouper, funcs):
    df = gr.get_frame(funcs)
    assert df.columns.to_list() == ["a", "b", "c"]


def test_get_frame_wrap_str(gr: StatsGrouper):
    df = gr.get_frame(["mean"], wrap="__{}__")
    values = df.to_numpy().flatten()
    assert all(x.startswith("=__AGGREGATE") for x in values)
    assert all(x.endswith(")__") for x in values)


def test_get_frame_wrap_dict(gr: StatsGrouper):
    df = gr.get_frame(["mean"], wrap={"a": "A{}A", "b": "B{}B"})
    assert all(x.startswith("=AAGGREGATE") for x in df["a"])
    assert all(x.endswith(")A") for x in df["a"])
    assert all(x.startswith("=BAGGREGATE") for x in df["b"])
    assert all(x.endswith(")B") for x in df["b"])
    assert all(x.startswith("=AGGREGATE") for x in df["c"])
    assert all(x.endswith(")") for x in df["c"])


def test_get_frame_offset(gr: StatsGrouper):
    df = gr.get_frame(["mean"]).reset_index()
    assert df["x"].iloc[0] == "=$C$4"
    assert df["y"].iloc[-1] == "=$D$16"
    assert df["a"].iloc[0] == "=AGGREGATE(1,7,$F$4:$F$7,$F$20:$F$23)"
    assert df["c"].iloc[-1] == "=AGGREGATE(1,7,$H$16:$H$19)"
