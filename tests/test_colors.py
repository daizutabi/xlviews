import pytest


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
    from xlviews.colors import rgb

    assert rgb(name) == color
    if isinstance(name, tuple):
        assert rgb(*name) == color


@pytest.mark.parametrize("name", ["invalid", (1, "x", "y")])
def test_rgb_error(name):
    from xlviews.colors import rgb

    with pytest.raises(ValueError, match="Invalid color format"):
        rgb(name)
