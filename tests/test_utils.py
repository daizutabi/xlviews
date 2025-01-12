import pytest
from pandas import DataFrame
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


def test_validate_list(sheet: Sheet):
    from xlviews.utils import add_validate_list

    rng = sheet.range("a1")
    add_validate_list(rng, [1, 2, 3], 2)
    assert rng.value == 2

    assert rng.api.Validation.Type == 3
    assert rng.api.Validation.Operator == 3
    assert rng.api.Validation.Formula1 == "1,2,3"
