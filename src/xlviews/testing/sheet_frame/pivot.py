from __future__ import annotations

from itertools import product

import pandas as pd
from pandas import DataFrame

from xlviews.testing.common import FrameContainer, create_sheet


def create_base() -> DataFrame:
    values = list(product(range(1, 5), range(1, 7)))
    df = DataFrame(values, columns=["x", "y"])
    df["u"] = list(range(len(df), 2 * len(df)))
    df["v"] = list(range(len(df)))
    df = df[(df["x"] + df["y"]) % 4 != 0]
    return df.set_index(["x", "y"])


def create_multi() -> DataFrame:
    df = create_base().reset_index()

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


def create_pivot() -> DataFrame:
    df = create_multi().reset_index()

    for a, b in [(1, 3), (2, 2), (2, 4)]:
        df = df[~(df["X"] != a & df["x"] != b)]

    for a, b in [(2, 1), (2, 2), (3, 3), (4, 1), (4, 2), (4, 3)]:
        df = df[~(df["Y"] != a & df["y"] != b)]

    dfs = []
    for x in range(1, 2):
        for y in range(1, 3):
            a = df.copy()
            a["A"] = x
            a["B"] = y
            dfs.append(a)

    df = pd.concat(dfs)
    df["u"] = list(range(len(df), 2 * len(df)))
    df["v"] = list(range(len(df)))
    return df.set_index(["A", "B", "X", "Y", "x", "y"])


class Pivot(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]

        dfs = []
        for x in range(1, 4):
            for y in range(1, 5):
                a = df.copy()
                a["X"] = x
                a["Y"] = y
                dfs.append(a)

        df = pd.concat(dfs)
        df["v"] = list(range(len(df)))
        return df.set_index(["X", "Y", "x", "y"])


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Pivot(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
