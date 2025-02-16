from __future__ import annotations

from itertools import product

import pandas as pd
from pandas import DataFrame

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing.common import FrameContainer, create_sheet


class Base(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])
        df["u"] = list(range(len(df), 2 * len(df)))
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]
        return df.set_index(["x", "y"])


class MultiIndex(FrameContainer):
    column: int = 17

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = Base.dataframe().reset_index()

        dfs = []
        for x in range(1, 4):
            for y in range(1, 5):
                a = df.copy()
                a["X"] = x
                a["Y"] = y
                dfs.append(a)

        df = pd.concat(dfs)
        df["u"] = list(range(len(df), 2 * len(df)))
        df["v"] = list(range(len(df)))
        return df.set_index(["X", "Y", "x", "y"])


class Pivot(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        df = MultiIndex.dataframe().reset_index()

        for a, b in [(1, 3), (2, 2), (2, 4)]:
            df = df[~((df["X"] == a) & (df["x"] == b))]

        for a, b in [(2, 1), (3, 3), (4, 1), (4, 3)]:
            df = df[~((df["Y"] == a) & (df["y"] == b))]

        dfs = []
        for x in range(1, 3):
            for y in range(1, 4):
                if x == 1 and y == 3:
                    continue

                a = df.copy()
                a["A"] = x
                a["B"] = y
                dfs.append(a)

        df = pd.concat(dfs)

        for a, b in [(2, 1)]:
            df = df[~((df["A"] == a) & (df["X"] == b))]

        for a, b in [(2, 3)]:
            df = df[~((df["B"] == a) & (df["Y"] == b))]

        df["u"] = (
            1000 * df["A"] + 300 * df["B"] + 100 * df["X"] + 10 * df["Y"] + df["x"]
        )
        df["v"] = list(range(len(df)))
        return df.set_index(["A", "B", "X", "Y", "x", "y"])


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    df = sf.pivot_table("u", "y", "x", formula=True)
    SheetFrame(2, 7, df).style()
    df = sf.pivot_table(None, "y", "x", formula=True)
    SheetFrame(10, 7, df).style().autofit()
    df = sf.pivot_table("u", None, "x", formula=True, aggfunc="mean")
    SheetFrame(19, 7, df).style().autofit()
    df = sf.pivot_table("v", "y", None, formula=True, aggfunc="max")
    SheetFrame(22, 7, df).style().autofit()

    sf.set_adjacent_column_width(1)
    fc = MultiIndex(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    df = sf.pivot_table("u", ["Y", "y"], ["X", "x"], formula=True)
    SheetFrame(2, 24, df).style()
    df = sf.pivot_table("v", "Y", ["X", "y", "x"], formula=True)
    SheetFrame(29, 24, df).style().autofit()
    df = sf.pivot_table(["u", "v"], ["Y", "y"], ["X", "x"], formula=True)
    SheetFrame(37, 24, df).style()
