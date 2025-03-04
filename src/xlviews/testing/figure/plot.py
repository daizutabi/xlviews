from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.figure.plot import Plot
from xlviews.testing.chart import Base
from xlviews.testing.common import create_sheet

if TYPE_CHECKING:
    from collections.abc import Hashable

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

    ax = Axes()
    data = sf.groupby("b").agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatterLines)
        .set(label="b={b}", marker=["o", "s"], color={"s": "red", "t": "blue"})
    )

    ax = Axes()
    data = sf.groupby(["b", "c"]).agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatterLines)
        .set(
            label=lambda x: f"{x['b']},{x['c']}",
            marker="b",
            color=("c", ["red", "green"]),
            size=10,
        )
    )

    def m(x: Hashable) -> str:
        if x == "s":
            return "o"
        return "^"

    ax = Axes(left=0)
    data = sf.groupby("b").agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatter)
        .set(label="{b}", marker=m, size=10)
    )

    def c(x: Hashable) -> str:
        if x == ("s", 100):
            return "red"
        return "blue"

    ax = Axes()
    data = sf.groupby(["b", "c"]).agg(include_sheetname=True)
    p = (
        Plot(ax, data)
        .add("x", "y", ChartType.xlXYScatter)
        .set(label="{b}_{c}", color=c, marker=("b", m), size=10)
    )
