from __future__ import annotations

from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.testing.common import create_sheet
from xlviews.testing.sheet_frame.base import MultiColumn

if __name__ == "__main__":
    sheet = create_sheet()
    fc = MultiColumn(sheet, column=2, style=True)
    sf = fc.sf

    ax = Axes(row=13, column=2, chart_type=ChartType.xlXYScatter)
    df = sf.groupby("s")  # .agg(include_sheetname=True)
    # for key, s in df.iterrows():
    #     print(key, s["x"], s["y"])
    #     ax.add_series(s["x"], s["y"], label=f"{key}")
    # x = sf.range("x")
    # y = sf.range("y")
    # label = sf.first_range("a")
    # ax.add_series(x, y, label=label)
    # ax.add_series(x.get_address(include_sheetname=True), y)
    # ax.chart.api[1].ChartTitle.Text = sheet_module["A1"].api
    # ax.chart.api[1].ChartTitle.Text = "=a1"

    # ax.xlabel = sf.range("x", -1)
    # ax.ylabel = sf.range("y", -1)

    # gr = sf.groupby(None)

    # gr = sf.groupby(["a", "b", "c"])
    # key = ("c", "t")
    # x = gr.range("x", key)
    # y = gr.range("y", key)
    # label = gr.first_range("b", key)
    # ax.add_series(x, y, label=label)

    # for x, y in zip(gr.ranges("x"), gr.ranges("y"), strict=True):
    #     ax.add_series(x, y)

    # list(gr.keys())

    # list(gr.first_ranges("a"))

    # ax.tight_layout()
    # ax.set_style()
    # ax.set_legend(loc=(0, 0))

    # g = data.groupby("a")
    # key = ("u",)
    # x = data.range("x", g[key])
    # y = data.range("y", g[key])
    # x
    # s = ax.add_series(
    #     x,
    #     y,
    #     # label=data.range("a", g["v"])[0],
    #     chart_type=ChartType.xlXYScatterLines,
    # )
    # ax.xlabel = data.range("x", 0)
    # ax.ylabel = data.range("y", 0)
    # ax.title = "=" + data.range("a", 0).get_address(include_sheetname=True)
