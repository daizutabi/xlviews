import pytest


def test_sns_color_palette():
    import seaborn as sns

    cp = sns.color_palette()
    assert len(cp) == 10


def test_color_palette():
    from xlviews.style import color_palette

    cp = color_palette(3)
    assert cp == [(31, 119, 180), (255, 127, 14), (44, 160, 44)]


def test_color_palette_large():
    from xlviews.style import color_palette

    cp = color_palette(20)
    assert cp[0] == (224, 37, 89)
    assert cp[-1] == (216, 40, 137)


def test_marker_palette():
    from xlviews.style import marker_palette

    mp = marker_palette(10)
    assert mp[0] == "o"
    assert mp[-2] == "*"
    assert mp[-1] == "o"


@pytest.mark.parametrize(
    ("name", "p"),
    [
        ("color", [(31, 119, 180), (255, 127, 14)]),
        ("marker", ["o", "^"]),
        ("none", [None, None]),
    ],
)
def test_palette_color(name: str, p):
    from xlviews.style import palette

    assert palette(name, 2) == p


def test_marker_style_int():
    from xlviews.style import get_marker_style

    assert get_marker_style(1) == 1


def test_line_style_int():
    from xlviews.style import get_line_style

    assert get_line_style(1) == 1
