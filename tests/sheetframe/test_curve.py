import string

import numpy as np
import pytest
from pandas import DataFrame, MultiIndex, Series
from xlwings import Sheet

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


def test_value(sf: SheetFrame):
    v = sf.cell.expand().options(ndim=2).value
    assert len(v) == 10
    assert v[2][:5] == ["r", 1, 1, 2, 2]
    assert v[3][:5] == ["i", "x", "y", "x", "y"]
    assert v[-1][:5] == [5, 5, 11, 17, 23]


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.cell.get_address() == "$B$2"
    assert sf.row == 2
    assert sf.column == 2
    assert sf.sheet.name == sheet_module.name
    assert sf.index_level == 1
    assert sf.columns_level == 4
    assert sf.columns_names == ["s", "t", "r", "i"]


def test_len(sf: SheetFrame):
    assert len(sf) == 6


def test_columns(sf: SheetFrame):
    columns = sf.columns
    assert columns[0] == ("s", "t", "r", "i")
    assert columns[1] == ("a", "c", 1, "x")
    assert columns[-1] == ("b", "d", 8, "y")


def test_value_columns(sf: SheetFrame):
    columns = sf.value_columns
    assert columns[0] == ("a", "c", 1, "x")
    assert columns[-1] == ("b", "d", 8, "y")


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == [("s", "t", "r", "i")]


@pytest.mark.parametrize(
    ("column", "index"),
    [
        (("s", "t", "r", "i"), 2),
        (("a", "d", 3, "x"), 7),
        ([("b", "c", 6, "y"), ("b", "d", 8, "x")], [14, 17]),
        ("s", 2),
        (["t", "i"], [3, 5]),
    ],
)
def test_index(sf: SheetFrame, column, index):
    assert sf.index(column) == index


@pytest.mark.parametrize("column", ["s", "t", "r", "i"])
def test_index_row(sf: SheetFrame, column):
    r = sf.index(column)
    assert sf.sheet.range(r, sf.column).value == column


@pytest.mark.parametrize(
    ("column", "index"),
    [
        ({"s": "a"}, (3, 10)),
        ({"s": "b", "t": "c"}, (11, 14)),
        ({"r": 5}, (11, 12)),
        ({"r": 7, "i": "y"}, (16, 16)),
    ],
)
def test_index_dict(sf: SheetFrame, column, index):
    assert sf.index_dict(column) == index
    assert sf.index(column) == index


def test_index_dict_error(sf: SheetFrame):
    with pytest.raises(NotImplementedError):
        sf.index_dict({"t": "c"})


def test_range_all(sf: SheetFrame):
    assert sf.range_all().get_address() == "$B$2:$R$11"
    assert sf.range().get_address() == "$B$2:$R$11"


@pytest.mark.parametrize(
    ("start", "end", "address"),
    [
        (0, None, "$B$6"),
        (50, None, "$B$50"),
        (15, 16, "$B$15:$B$16"),
    ],
)
def test_range_index(sf: SheetFrame, start: int | None, end, address):
    assert sf.range_index(start, end).get_address() == address
    assert sf.range("index", start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        (("a", "c", 1, "y"), False, None, "$D$2:$D$11"),
        (("a", "d", 4, "y"), -1, None, "$J$2:$J$5"),
        (("a", "d", 3, "x"), 0, None, "$G$6"),
        (("b", "d", 7, "x"), None, None, "$O$6:$O$11"),
        (("a", "d", 3, "y"), 100, None, "$H$100"),
    ],
)
def test_range_column(sf: SheetFrame, column, start, end, address):
    assert sf.range_column(column, start, end).get_address() == address
    assert sf.range(column, start, end).get_address() == address


@pytest.mark.parametrize(
    ("column", "start", "end", "address"),
    [
        (("a", "d", 3, "y"), [(3, 5), (7, 8)], None, "$H$3:$H$5,$H$7:$H$8"),
    ],
)
def test_range_start_list(sf: SheetFrame, column, start, end, address):
    assert sf.range(column, start, end).get_address() == address


def test_range_column_error(sf: SheetFrame):
    with pytest.raises(NotImplementedError):
        sf.range_column("s")

    with pytest.raises(NotImplementedError):
        sf.range("t")


def test_getitem_tuple(sf: SheetFrame):
    s = sf[("a", "d", 4, "x")]
    assert isinstance(s, Series)
    assert s.name == ("a", "d", 4, "x")
    np.testing.assert_array_equal(s, [36, 37, 38, 39, 40, 41])
    s = sf["a", "d", 4, "x"]
    np.testing.assert_array_equal(s, [36, 37, 38, 39, 40, 41])


def test_getitem_list(sf: SheetFrame):
    df = sf[[("a", "d", 4, "x"), ("a", "d", 4, "y")]]
    assert isinstance(df, DataFrame)
    assert df.columns.to_list() == [("a", "d", 4, "x"), ("a", "d", 4, "y")]
    x = np.arange(36, 48).reshape(2, 6).T
    np.testing.assert_array_equal(df, x)


@pytest.mark.parametrize(
    ("kwargs", "sel"),
    [
        ({"s": "a"}, [True] * 8 + [False] * 8),
        ({"i": "x"}, [True, False] * 8),
        ({"s": "a", "i": "x"}, [True, False] * 4 + [False] * 8),
        ({"r": (3, 6)}, [False] * 4 + [True] * 8 + [False] * 4),
        ({"r": [1, 8]}, [True] * 2 + [False] * 12 + [True] * 2),
    ],
)
def test_select(sf: SheetFrame, kwargs, sel):
    x = sf.select(**kwargs)
    np.testing.assert_array_equal(x, sel)


def test_ranges(sf: SheetFrame):
    for rng, i in zip(sf.ranges(), range(2, 18), strict=True):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


def test_ranges_sel(sf: SheetFrame):
    sel = [False] * 8 + [True] * 8
    for rng, i in zip(sf.ranges(sel), range(10, 18), strict=True):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


def test_ranges_kwargs(sf: SheetFrame):
    for rng, i in zip(sf.ranges(i="y"), range(3, 18, 2), strict=True):
        c = string.ascii_uppercase[i]
        assert rng.get_address() == f"${c}$6:${c}$11"


@pytest.mark.parametrize(
    ("by", "result"),
    [
        ("s", {("a",): [(3, 10)], ("b",): [(11, 18)]}),
        (
            ["s", "t"],
            {
                ("a", "c"): [(3, 6)],
                ("a", "d"): [(7, 10)],
                ("b", "c"): [(11, 14)],
                ("b", "d"): [(15, 18)],
            },
        ),
        (
            ["s", "i"],
            {
                ("a", "x"): [(3, 3), (5, 5), (7, 7), (9, 9)],
                ("a", "y"): [(4, 4), (6, 6), (8, 8), (10, 10)],
                ("b", "x"): [(11, 11), (13, 13), (15, 15), (17, 17)],
                ("b", "y"): [(12, 12), (14, 14), (16, 16), (18, 18)],
            },
        ),
        (None, {(): [(3, 18)]}),
    ],
)
def test_groupby(sf: SheetFrame, by, result):
    g = sf.groupby(by)
    assert len(g) == len(result)
    for k, v in g.items():
        assert result[k] == v
