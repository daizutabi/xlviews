if __name__ == "__main__":
    from itertools import product

    import pandas as pd
    import xlwings as xw

    from xlviews.dataframes.heat_frame import HeatFrame
    from xlviews.dataframes.sheet_frame import SheetFrame

    for app in xw.apps:
        app.quit()

    book = xw.Book()
    sheet = book.sheets.add()

    values = list(product(range(1, 5), range(1, 4)))
    df = pd.DataFrame(values, columns=["x", "y"])
    df["v"] = list(range(len(df)))
    df = df.set_index(["x", "y"])
    sf = SheetFrame(2, 2, data=df, index=True)

    data = sf.get_address(["v"], formula=True)
    hf = HeatFrame(2, 6, data=data, x="x", y="y", value="v")

    hf.range()
