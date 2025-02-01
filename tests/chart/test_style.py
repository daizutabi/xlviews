def test_marker_palette():
    from xlviews.chart.style import marker_palette

    mp = marker_palette(10)
    assert mp[0] == "o"
    assert mp[-2] == "*"
    assert mp[-1] == "o"


def test_marker_style_int():
    from xlviews.chart.style import get_marker_style

    assert get_marker_style(1) == 1


def test_line_style_int():
    from xlviews.chart.style import get_line_style

    assert get_line_style(1) == 1
