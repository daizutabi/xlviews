from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
from pandas import DataFrame

from xlviews.dataframes.stats_frame import StatsFrame
from xlviews.testing.common import FrameContainer, create_sheet

if TYPE_CHECKING:
    from xlwings import Sheet


class Parent(FrameContainer):
    row: int = 3
    column: int = 3

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "x": ["a"] * 8 + ["b"] * 8 + ["a"] * 4,
                "y": (["c"] * 4 + ["d"] * 4) * 2 + ["c"] * 4,
                "z": range(1, 21),
                "a": range(20),
                "b": list(range(10)) + list(range(0, 30, 3)),
                "c": list(range(20, 40, 2)) + list(range(0, 20, 2)),
            },
        )
        df = df.set_index(["x", "y", "z"])
        df.iloc[[4, -1], 0] = np.nan
        df.iloc[[3, 6, 9], -1] = np.nan
        return df


def main(sheet: Sheet) -> None:
    fc = Parent(sheet, 3, 3, style=True, table=True)
    sf_parent = fc.sf

    funcs = ["count", "max", "median", "soa"]
    StatsFrame(sf_parent, funcs, by=":y", table=True)


if __name__ == "__main__":
    sheet = create_sheet()
    main(sheet)
