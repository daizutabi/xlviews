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

    x, y = sf.agg(include_sheetname=True)
    ax = Axes(100, 20, chart_type=ChartType.xlXYScatter)
    ax.add_series(x, y, label="label")
    ax.set(
        xlabel="xlabel",
        ylabel="ylabel",
        title="title",
        xticks=(0, 10, 2),
        legend=(1, -1),
    )

    df = sf.groupby("b").agg(include_sheetname=True)
    print(df.index)
    for key, (x, y) in df.iterrows():
        print(key, x, y)

    df = DataFrame([[1, 2], [4, 5]], columns=["a", "b"])
    a = df.groupby(["a"]).agg("mean")
    print(a)
    print(a.index)
    # ax = Axes(chart_type=ChartType.xlXYScatterLines)
    # ax.add_series(x, y, label="label")
    # ax.set(
    #     xlabel="xlabel",
    #     ylabel="ylabel",
    #     title="title",
    # )
    # x = sf.range("x")
    # y = sf.range("y")
    # label = sf.first_range("a")
    # ax.add_series(x.get_address(include_sheetname=True), y)
    #     fc.sf.set_adjacent_column_width(1)
    #     fc.sf.number_format(b="0.0")
    #     DistFrame(fc.sf, ["a", "b"], by=["x", "y"])

    # if __name__ == "__main__":
    #     import numpy as np
    #     import xlwings as xw
    #     from pandas import DataFrame, MultiIndex
    #     from xlwings.constants import ChartType

    #     from xlviews.chart.axes import Axes
    #     from xlviews.dataframes.sheet_frame import SheetFrame

    #     for app in xw.apps:
    #         app.quit()

    #     book = xw.Book()
    #     sheet_module = book.sheets.add()

    # a = ["a"] * 8 + ["b"] * 8
    # b = (["c"] * 4 + ["d"] * 4) * 2
    # c = np.repeat(range(1, 9), 2)
    # d = ["x", "y"] * 8
    # df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
    # df.columns = MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
    # sf = SheetFrame(2, 2, df, sheet_module)

    # len(list(sf.iter_ranges()))

    # columns = sf.value_columns
    # DataFrame(columns, columns=sf.columns_names)

    # df.melt()

    # a = ["c"] * 10
    # b = ["s"] * 5 + ["t"] * 5
    # c = ([100] * 2 + [200] * 3) * 2
    # x = list(range(10))
    # y = list(range(10, 20))
    # df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    # df = df.set_index(["a", "b", "c"])
    # sf = SheetFrame(2, 2, df, sheet_module)

    # ax = Axes(left=200, chart_type=ChartType.xlXYScatter)
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
