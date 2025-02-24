from __future__ import annotations

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.chart.plot import plot
from xlviews.testing.common import create_sheet

from .base import Base

if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    data = sf.agg(include_sheetname=True)
    ax = Axes(2, 8, chart_type=ChartType.xlXYScatter)
    plot(ax, data, "x", "y")
