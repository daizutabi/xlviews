if __name__ == "__main__":
    import xlwings as xw
    from pandas import DataFrame
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.common import quit_apps
    from xlviews.plot import _plot
    from xlviews.sheetframe import SheetFrame

    quit_apps()
    book = xw.Book()
    sheet_module = book.sheets.add()
    x = (["u"] * 2 + ["v"] * 3) * 2
    df = DataFrame({"a": x, "x": list(range(10)), "y": list(range(10, 20))})
    data = SheetFrame(sheet_module, 2, 2, data=df, index=False)

    ax = Axes(left=200, chart_type=ChartType.xlXYScatter)
    g = data.groupby("a")
    x = data.range("x", g["v"])
    y = data.range("y", g["v"])
    s = ax.add_series(
        x,
        y,
        label=data.range("a", g["v"])[0],
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
