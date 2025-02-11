from __future__ import annotations

from itertools import product
from typing import TYPE_CHECKING

import pandas as pd
from pandas import DataFrame

from xlviews.dataframes.heat_frame import pivot_table
from xlviews.testing.common import FrameContainer, create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer

if TYPE_CHECKING:
    from xlviews.dataframes.sheet_frame import SheetFrame


class BaseParent(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = list(product(range(1, 5), range(1, 7)))
        df = DataFrame(values, columns=["x", "y"])
        df["v"] = list(range(len(df)))
        df = df[(df["x"] + df["y"]) % 4 != 0]

        return df.set_index(["x", "y"])


class Base(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        data = sf.get_address(["v"], formula=True)
        return pivot_table(data, "v", "x", "y")

    def init(self) -> None:
        super().init()
        self.sf.set_label("v")


class MultiIndexParent(FrameContainer):
    column: int = 14

    @classmethod
    def dataframe(cls) -> DataFrame:
        base = BaseParent.dataframe().reset_index()
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


class MultiIndex(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        data = sf.get_address(["v"], formula=True)
        return pivot_table(data, "v", ["X", "x"], ["Y", "y"])


if __name__ == "__main__":
    sheet = create_sheet()

    fc = BaseParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    fc = Base(sf)
    fc.sf.set_adjacent_column_width(1)

    fc = MultiIndexParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    fc = MultiIndex(sf)
    fc.sf.set_adjacent_column_width(1)
