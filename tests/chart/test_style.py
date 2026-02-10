from __future__ import annotations

from xlviews.chart.style import get_line_style, get_marker_style


def test_marker_style_int():
    assert get_marker_style(1) == 1


def test_line_style_int():
    assert get_line_style(1) == 1
