if __name__ == "__main__":
    import numpy as np
    import xlwings as xw
    from pandas import DataFrame, MultiIndex
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.sheetframe import SheetFrame

    for app in xw.apps:
        app.quit()

    book = xw.Book()
    sheet_module = book.sheets.add()

    a = ["a"] * 8 + ["b"] * 8
    b = (["c"] * 4 + ["d"] * 4) * 2
    c = np.repeat(range(1, 9), 2)
    d = ["x", "y"] * 8
    df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
    df.columns = MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
    sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet_module)

    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet_module)

    ax = Axes(left=200, chart_type=ChartType.xlXYScatter)
    x = sf.range("x")
    y = sf.range("y")
    label = sf.first_range("a")
    ax.add_series(x, y, label=label)

    ax.xlabel = sf.range("x", -1)
    ax.ylabel = sf.range("y", -1)

    gr = sf.grouper(None)

    gr = sf.grouper(["a", "b", "c"])
    key = ("c", "t")
    x = gr.range("x", key)
    y = gr.range("y", key)
    label = gr.first_range("b", key)
    ax.add_series(x, y, label=label)

    for x, y in zip(gr.ranges("x"), gr.ranges("y"), strict=True):
        ax.add_series(x, y)

    list(gr.keys())

    list(gr.first_ranges("a"))

    ax.tight_layout()
    ax.set_style()
    ax.set_legend(loc=(0, 0))

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
