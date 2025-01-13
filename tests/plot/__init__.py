if __name__ == "__main__":
    import numpy as np
    import xlwings as xw
    from pandas import DataFrame, MultiIndex
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.grouper import Grouper
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
    gr = Grouper(sf, ["s", "t"])

    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet_module)
    gr = Grouper(sf, ["b", "c"])

    ax = Axes(left=200, chart_type=ChartType.xlXYScatter)
    x = sf.range("x", -1)
    y = sf.range("y", -1)
    label = sf.range("a")
    ax.add_series(x, y, label=label)

    gr = sf.grouper(["a", "b"])

    gr.range("x", ("c", "s"))
    gr.first_range("a", ("c", "s"))

    g = data.groupby("a")
    key = ("u",)
    x = data.range("x", g[key])
    y = data.range("y", g[key])
    x
    s = ax.add_series(
        x,
        y,
        # label=data.range("a", g["v"])[0],
        chart_type=ChartType.xlXYScatterLines,
    )
    ax.xlabel = data.range("x", 0)
    ax.ylabel = data.range("y", 0)
    ax.title = "=" + data.range("a", 0).get_address(include_sheetname=True)
    ax.tight_layout()
    ax.set_plot_area_style()
    ax.set_legend(loc=(0, 0))
    s.set(marker="s", size=5, line="--", color="blue", line_weight=2, alpha=0.3)
    s.x = [1, 2, 3, 2]
    s.y = [1, 2, 3, 5]
    s.x

    df.groupby("a")
