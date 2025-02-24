from __future__ import annotations

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.plot.core import plot_series
from xlviews.testing.common import create_sheet

from .base import Base

if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    s = sf.agg(include_sheetname=True)
    ax = Axes(2, 8, chart_type=ChartType.xlXYScatter)
    s.name = ("a", "b")
    plot_series(ax, s, "x", "y", label="a{0}b{1}c")

    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 10, 2),
        legend=(1, -1),
    )
