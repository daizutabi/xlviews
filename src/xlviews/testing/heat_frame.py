from __future__ import annotations

from itertools import product

import numpy as np
import pandas as pd
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.dataframes.stats_frame import StatsFrame
from xlviews.testing.common import FrameContainer, create_sheet


class HeatContainer(FrameContainer):
    src: DataFrame
    include_sheetname: bool = False

    def init(self) -> None:
        columns = self.sf.value_columns
        self.src = self.sf.get_address(columns, formula=True)
        self.sf.sheet.range("a1").get_address()


class Base(HeatContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]

        return df.set_index(["x", "y"])


class Agg(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        base = Base.dataframe()
        dfs = []
        for x in range(3):
            for y in range(4):
                a = base.copy()
                a["X"] = x
                a["Y"] = y
                dfs.append(a)

        df = pd.concat(dfs)
        df["v"] = list(range(len(df)))
        return df.set_index(["X", "Y", "x", "y"])


class Parent(FrameContainer):
    row: int = 2
    column: int = 2

    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])

        dfs = []
        for x in range(3):
            for y in range(2):
                a = df.copy()
                a["X"] = x
                a["Y"] = y
                dfs.append(a)

        df = pd.concat(dfs)
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]

        return df.set_index(["X", "Y", "x", "y"])

    # sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet_module)
    # data = sf.get_address(["v"], formula=True)

    # return HeatFrame(2, 6, data=data, x="x", y="y", value="v", sheet=sheet_module)


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Base(sheet, style=True)
    sf = fc.sf
    data = sf.get_address(["v"], formula=True)
    sf.set_adjacent_column_width(1)
    HeatFrame(2, 6, data=data, x="x", y="y", value="v", sheet=sheet)
