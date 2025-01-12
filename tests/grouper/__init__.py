if __name__ == "__main__":
    import numpy as np
    import xlwings as xw
    from pandas import DataFrame, MultiIndex

    from xlviews.common import quit_apps
    from xlviews.grouper import Grouper
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
    gr = Grouper(sf, ["b", "c"])

    a = ["a"] * 8 + ["b"] * 8
    b = (["c"] * 4 + ["d"] * 4) * 2
    c = np.repeat(range(1, 9), 2)
    d = ["x", "y"] * 8
    df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
    df.columns = MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
    sf = SheetFrame(sheet_module, 2, 2, data=df, index=True)

    gr = Grouper(sf, ["s", "t"])
