if __name__ == "__main__":
    import xlwings as xw
    from pandas import DataFrame

    from xlviews.common import quit_apps
    from xlviews.sheetframe import SheetFrame

    quit_apps()
    book = xw.Book()
    sheet_module = book.sheets.add()
    a = ["c"] * 10
    b = ["s"] * 5 + ["t"] * 5
    c = ([100] * 2 + [200] * 3) * 2
    x = list(range(10))
    y = list(range(10, 20))
    df = DataFrame({"a": a, "b": b, "c": c, "x": x, "y": y})
    df = df.set_index(["a", "b", "c"])
    sf = SheetFrame(sheet_module, 2, 2, data=df, index=True)
