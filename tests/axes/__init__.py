if __name__ == "__main__":
    import xlwings as xw
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.common import quit_apps
    from xlviews.style import set_series_style

    quit_apps()
    book = xw.Book()
    sheet_module = book.sheets.add()

    ct = ChartType.xlXYScatterLines
    ax = Axes(300, 10, chart_type=ct, sheet=sheet_module)
    x = sheet_module["B2:B11"]
    y = sheet_module["C2:C11"]
    x.options(transpose=True).value = list(range(10))
    y.options(transpose=True).value = list(range(10, 20))

    s = ax.add_series(x, y, label="a")
    set_series_style(s, marker="d", size=10, line="--", color="red", alpha=0.5)
