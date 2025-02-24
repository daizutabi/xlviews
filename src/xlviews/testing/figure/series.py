from __future__ import annotations

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.figure.plot import Plot
from xlviews.testing.chart import Base
from xlviews.testing.common import create_sheet

if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    ax = Axes(2, 8, chart_type=ChartType.xlXYScatter)

    # data = sf.agg(include_sheetname=True)
    data = sf.groupby("b").agg(include_sheetname=True)
    # data = sf.groupby(["b", "c"]).agg(include_sheetname=True)

    p = Plot(ax, data).add("x", "y")
    p.set(label=lambda x: f"--{x['b']}--")

    # s.name = ("a", "b")
    # print(s)
    # a = s.to_frame().T
    # print(a.index.names)
    # for i, r in a.iterrows():
    #     print(i, r)
    # plot_series(ax, s, "x", "y", label="a{0}b{1}c", color="red")
    # plot_series(ax, s, "x", "y", label=lambda x: str(x), color=lambda x: "blue")
    # plot_series(ax, s, "x", "y", label={("a", "b"): "d"}, color={("a", "b"): "green"})

    # ax.set(
    #     xlabel="xlabel",
    #     ylabel="ylabel",
    #     title="title",
    #     xticks=(0, 10, 2),
    #     legend=(1, -1),
    # )
