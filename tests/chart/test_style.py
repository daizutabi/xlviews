from itertools import islice


def test_marker_style_int():
    from xlviews.chart.style import get_marker_style

    assert get_marker_style(1) == 1


def test_line_style_int():
    from xlviews.chart.style import get_line_style

    assert get_line_style(1) == 1


def test_cycle_colors():
    from xlviews.chart.style import cycle_colors

    x = list(islice(cycle_colors(), 3))
    assert x == ["#1f77b4", "#ff7f0e", "#2ca02c"]


def test_cycle_colors_skips():
    from xlviews.chart.style import cycle_colors

    x = list(islice(cycle_colors(["#ff7f0e"]), 3))
    assert x == ["#1f77b4", "#2ca02c", "#d62728"]


def test_cycle_markers():
    from xlviews.chart.style import cycle_markers

    x = list(islice(cycle_markers(), 3))
    assert x == ["o", "^", "s"]


def test_cycle_markers_skips():
    from xlviews.chart.style import cycle_markers

    x = list(islice(cycle_markers(["o"]), 3))
    assert x == ["^", "s", "d"]
