import pytest


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
