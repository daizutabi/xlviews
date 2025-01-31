import numpy as np
import pandas as pd
import xlwings as xw
from pandas import DataFrame

from xlviews.frame import SheetFrame


def get_unit(a: int, b: int, o: int) -> DataFrame:
    x = np.repeat(range(1, 5), 9)
    y = np.repeat(list(range(1, 4)) * 4, 3)
    u = range(o, len(x) + o)
    v = range(len(x) + o, 2 * len(x) + o)
    df = DataFrame({"a": a, "b": b, "x": x, "y": y, "u": u, "v": v})
    return df.query("not (x == 4 and y == 3)").copy()


def get_data() -> DataFrame:
    df = pd.concat([get_unit(1, 1, 0), get_unit(1, 2, 100), get_unit(2, 2, 200)])
    return df.set_index(["a", "b", "x", "y"])


def main():
    book = xw.Book()
    sheet = book.sheets.add()

    df = get_data()
    print(df)
    print(df.loc[(1, 1)].pivot_table(index="y", columns="x", values="u"))
    sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet)

    df = sf.groupby(["a", "b", "x", "y"]).agg("mean", formula=True)
    sf = SheetFrame(2, 9, data=df, index=True, sheet=sheet)

    a = df.loc[(1, 1)].reset_index()
    x = a.pivot(index="y", columns="x", values="u")
    print(x)


if __name__ == "__main__":
    for app in xw.apps:
        app.quit()

    main()
