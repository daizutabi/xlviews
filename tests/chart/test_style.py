def test_marker_style_int():
    from xlviews.chart.style import get_marker_style

    assert get_marker_style(1) == 1


def test_line_style_int():
    from xlviews.chart.style import get_line_style

    assert get_line_style(1) == 1
