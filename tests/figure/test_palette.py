from itertools import islice

import pytest
from pandas import DataFrame


@pytest.fixture(scope="module")
def df():
    return DataFrame(
        {
            "a": [1, 1, 2, 2, 3, 3],
            "b": [4, 4, 5, 6, 7, 7],
            "c": [10, 10, 10, 11, 11, 11],
        },
    )


@pytest.mark.parametrize(
    ("columns", "index"),
    [
        (["a"], {(1,): 0, (2,): 1, (3,): 2}),
        (["b"], {(4,): 0, (5,): 1, (6,): 2, (7,): 3}),
        (["c"], {(10,): 0, (11,): 1}),
        (["a", "b"], {(1, 4): 0, (2, 5): 1, (2, 6): 2, (3, 7): 3}),
        (["a", "c"], {(1, 10): 0, (2, 10): 1, (2, 11): 2, (3, 11): 3}),
        (["b", "c"], {(4, 10): 0, (5, 10): 1, (6, 11): 2, (7, 11): 3}),
    ],
)
def test_get_index(df: DataFrame, columns, index):
    from xlviews.figure.palette import get_index

    assert get_index(df[columns]) == index


@pytest.mark.parametrize(
    ("columns", "default", "index"),
    [
        (["a"], [3, 2], {(1,): 2, (2,): 1, (3,): 0}),
        (["b"], [6, 7], {(4,): 2, (5,): 3, (6,): 0, (7,): 1}),
        (["c"], [12, 11], {(10,): 1, (11,): 0}),
        (["a", "b"], [[2, 5], [3, 7]], {(1, 4): 2, (2, 5): 0, (2, 6): 3, (3, 7): 1}),
        (
            ["a", "c"],
            [[2, 10], (3, 11)],
            {(1, 10): 2, (2, 10): 0, (2, 11): 3, (3, 11): 1},
        ),
        (["b", "c"], [], {(4, 10): 0, (5, 10): 1, (6, 11): 2, (7, 11): 3}),
    ],
)
def test_get_index_default(df: DataFrame, columns, default, index):
    from xlviews.figure.palette import get_index

    assert get_index(df[columns], default) == index


def test_cycle_colors():
    from xlviews.figure.palette import cycle_colors

    x = list(islice(cycle_colors(), 3))
    assert x == ["#1f77b4", "#ff7f0e", "#2ca02c"]


def test_cycle_colors_skips():
    from xlviews.figure.palette import cycle_colors

    x = list(islice(cycle_colors(["#ff7f0e"]), 3))
    assert x == ["#1f77b4", "#2ca02c", "#d62728"]


def test_cycle_markers():
    from xlviews.figure.palette import cycle_markers

    x = list(islice(cycle_markers(), 3))
    assert x == ["o", "^", "s"]


def test_cycle_markers_skips():
    from xlviews.figure.palette import cycle_markers

    x = list(islice(cycle_markers(["o"]), 3))
    assert x == ["^", "s", "d"]


@pytest.mark.parametrize(("key", "value"), [(1, "o"), (2, "^"), (3, "s")])
def test_marker_palette(df: DataFrame, key, value):
    from xlviews.figure.palette import MarkerPalette

    p = MarkerPalette(df, "a")
    assert p[key] == value


@pytest.mark.parametrize(("key", "value"), [(1, "^"), (2, "o"), (3, "x")])
def test_marker_palette_default(df: DataFrame, key, value):
    from xlviews.figure.palette import MarkerPalette

    p = MarkerPalette(df, "a", {2: "o", 3: "x"})
    assert p[key] == value


@pytest.mark.parametrize(
    ("key", "value"),
    [((1, 4), "o"), ((2, 5), "^"), ((2, 6), "s")],
)
def test_marker_palette_multi(df: DataFrame, key, value):
    from xlviews.figure.palette import MarkerPalette

    p = MarkerPalette(df, ["a", "b"])
    assert p[key] == value


@pytest.mark.parametrize(("key", "value"), [(4, "#1f77b4"), (5, "red"), (6, "blue")])
def test_color_palette(df: DataFrame, key, value):
    from xlviews.figure.palette import ColorPalette

    p = ColorPalette(df, "b", {5: "red", 6: "blue"})
    assert p[key] == value
