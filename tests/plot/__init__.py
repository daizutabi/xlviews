if __name__ == "__main__":
    import xlwings as xw
    from pandas import DataFrame

    from xlviews.axes import Axes
    from xlviews.common import quit_apps
    from xlviews.sheetframe import SheetFrame

    quit_apps()
    book = xw.Book()
    sheet_module = book.sheets.add()
    df = DataFrame({"x": list(range(10)), "y": list(range(10, 20))})
    data = SheetFrame(sheet_module, 2, 2, data=df, index=False)
    ax = Axes()
