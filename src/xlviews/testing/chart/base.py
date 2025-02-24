from __future__ import annotations

from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.testing.common import FrameContainer, create_sheet


class Base(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        a = ["c"] * 10
        b = ["s"] * 5 + ["t"] * 5
        c = ([100] * 2 + [200] * 3) * 2
        x = list(range(10))
        y = list(range(10, 20))
        df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
        return df.set_index(["a", "b", "c"])


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    data = sf.agg(include_sheetname=True)
    ax = Axes(2, 8, chart_type=ChartType.xlXYScatter)
    s = ax.add_series(data["x"], data["y"])
    s.set("o", color="red", alpha=0.6)

    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 10, 2),
        legend=(1, -1),
    )

    ax = Axes(chart_type=ChartType.xlXYScatterLines)
    df = sf.groupby("b").agg(include_sheetname=True)
    for key, s in df.iterrows():
        x = ax.add_series(s["x"], s["y"], label=f"{key}")
        x.set("--", color="red", marker="o")

    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 10, 5),
        legend=(-1, -1),
    )

    ax = Axes(chart_type=ChartType.xlXYScatterLinesNoMarkers)
    df = sf.groupby(["b", "c"]).agg(include_sheetname=True)
    for key, s in df.iterrows():
        x = ax.add_series(s["x"], s["y"], label=f"{key[0]}_{key[1]}")  # type: ignore
        x.set(marker="o")

    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 20, 4),
        yticks=(0, 20, 4),
        legend=(1, 1),
    )
