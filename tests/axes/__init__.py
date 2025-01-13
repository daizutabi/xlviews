if __name__ == "__main__":
    import xlwings as xw
    from xlwings.constants import ChartType

    from xlviews.axes import Axes
    from xlviews.decorators import quit_apps

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
