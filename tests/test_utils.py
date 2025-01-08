import pytest
from pandas import DataFrame, Series
from xlwings import Sheet


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("ChartType.xlXYScatter", -4169),
        ("BordersIndex.EdgeTop", 8),
        ("Bottom", -4107),
        ("Center", -4108),
        ("Left", -4131),
        ("None", -4142),
        ("Right", -4152),
        ("Top", -4160),
    ],
)
def test_constant(name: str, value: int):
    from xlviews.utils import constant

    assert constant(name) == value
    assert constant(*name.split(".")) == value


@pytest.mark.parametrize("n", [1, 10, 100, 1000, 10000])
def test_column_name(sheet_module: Sheet, n):
    from xlviews.utils import int_to_column_name

    rng = sheet_module.range(1, n)
    assert rng.get_address().split("$")[1] == int_to_column_name(n)


@pytest.mark.parametrize(
    ("name", "color"),
    [
        ("black", 0),
        ("red", 2**8 - 1),
        ("green", 32768),
        ("blue", 256 * 256 * 255),
        ("white", 2**24 - 1),
        ("aqua", 16776960),
        ("gray", 8421504),
        ("indigo", 8519755),
        ("lime", 65280),
        ("pink", 13353215),
        ((10, 10, 10), 10 + 10 * 256 + 10 * 256 * 256),
        (100, 100),
    ],
)
def test_rgb(name, color):
    from xlviews.utils import rgb

    assert rgb(name) == color
    if isinstance(name, tuple):
        assert rgb(*name) == color


@pytest.mark.parametrize("name", ["invalid", (1, "x", "y")])
def test_rgb_error(name):
    from xlviews.utils import rgb

    with pytest.raises(ValueError, match="Invalid color format"):
        rgb(name)


@pytest.mark.parametrize(
    ("columns", "lst"),
    [
        ("B", ["B"]),
        (["A", "C"], ["A", "C"]),
        (":B", ["A", "B"]),
        (["::B", "C"], ["A", "C"]),
    ],
)
def test_iter_columns(columns, lst):
    from xlviews.utils import iter_columns

    df = DataFrame([[1, 2, 3], [4, 5, 6], [7, 8, 9]], columns=["A", "B", "C"])
    assert list(iter_columns(df, columns)) == lst


@pytest.mark.parametrize("func", [lambda x: x, Series])
def test_array_index(func):
    from xlviews.utils import array_index

    values = [1, 1, 2, 2, 2, 3, 3, 1, 1, 2, 2, 3, 3]
    index = array_index(func(values))
    assert index[1] == [[0, 1], [7, 8]]
    assert index[2] == [[2, 4], [9, 10]]
    assert index[3] == [[5, 6], [11, 12]]


@pytest.mark.parametrize("func", [lambda x: x, DataFrame])
def test_array_index_list(func):
    from xlviews.utils import array_index

    values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
    index = array_index(func(values))
    assert index[(1, 2)] == [[0, 1], [4, 4]]
    assert index[(3, 4)] == [[2, 3], [5, 6]]


@pytest.mark.parametrize("func", [lambda x: x, DataFrame])
def test_array_index_sel(func):
    from xlviews.utils import array_index

    values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
    sel = [True, False, True, False, True, False, True]
    index = array_index(func(values), sel=sel)
    assert index[(1, 2)] == [[0, 0], [4, 4]]
    assert index[(3, 4)] == [[2, 2], [6, 6]]


def test_array_index_empty():
    from xlviews.utils import array_index

    assert not array_index([])


def test_validate_list(sheet: Sheet):
    from xlviews.utils import add_validate_list

    rng = sheet.range("a1")
    add_validate_list(rng, [1, 2, 3], 2)
    assert rng.value == 2

    assert rng.api.Validation.Type == 3
    assert rng.api.Validation.Operator == 3
    assert rng.api.Validation.Formula1 == "1,2,3"
