from __future__ import annotations

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.figure.plot import plot_series
from xlviews.testing.chart import Base
from xlviews.testing.common import create_sheet

if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    s = sf.agg(include_sheetname=True)
    s.name = ("a", "b")
    print(s)
    a = s.to_frame().T
    print(a.index.names)
    for i, r in a.iterrows():
        print(i, r)
    ax = Axes(2, 8, chart_type=ChartType.xlXYScatter)
    plot_series(ax, s, "x", "y", label="a{0}b{1}c", color="red")
    plot_series(ax, s, "x", "y", label=lambda x: str(x), color=lambda x: "blue")
    plot_series(ax, s, "x", "y", label={("a", "b"): "d"}, color={("a", "b"): "green"})

    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 10, 2),
        legend=(1, -1),
    )
