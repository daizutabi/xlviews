from __future__ import annotations

from itertools import product

import pandas as pd
from pandas import DataFrame

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import FrameContainer, create_sheet


class Base(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]

        return df.set_index(["x", "y"])


class MultiIndex(FrameContainer):
    column: int = 14

    @classmethod
    def dataframe(cls) -> DataFrame:
        base = Base.dataframe().reset_index()
        dfs = []
        for x in range(1, 4):
            for y in range(1, 5):
                a = base.copy()
                a["X"] = x
                a["Y"] = y
                dfs.append(a)

        df = pd.concat(dfs)
        df["v"] = list(range(len(df)))
        return df.set_index(["X", "Y", "x", "y"])


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Base(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    data = sf.get_address(["v"], formula=True)
    sf = HeatFrame(2, 6, data)
    sf.set_adjacent_column_width(1)

    fc = MultiIndex(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    sf = HeatFrame(2, 20, sf, "v", ["X", "x"], ["Y", "y"])
