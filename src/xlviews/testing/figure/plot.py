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

    ax = Axes(2, 8)
    data = sf.agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatter)
        .set(label="abc", marker="o", color="blue", alpha=0.6)
    )

    ax = Axes(chart_type=ChartType.xlXYScatterLines)
    data = sf.groupby("b").agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y")
        .set(label="b={b}", marker=["o", "s"], color={"s": "red", "t": "blue"})
    )

    ax = Axes(chart_type=ChartType.xlXYScatterLines)
    data = sf.groupby(["b", "c"]).agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y")
        .set(
            label=lambda x: f"{x['b']},{x['c']}",
            marker="b",
            color=("c", ["red", "green"]),
            size=10,
        )
    )
