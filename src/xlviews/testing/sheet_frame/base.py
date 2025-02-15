from __future__ import annotations

import numpy as np
import pandas as pd
from pandas import DataFrame

from xlviews.testing.common import FrameContainer, create_sheet


class NoIndex(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        values = {"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]}
        return DataFrame(values)


class Index(FrameContainer):
    row: int = 8
    column: int = 2

    @classmethod
    def dataframe(cls) -> DataFrame:
        values = {"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]}
        index = ["x", "x", "y", "y"]
        df = DataFrame(values, index=index)
        df.index.name = "name"
        return df


class MultiIndex(FrameContainer):
    row: int = 2
    column: int = 6

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "x": [1, 1, 1, 1, 2, 2, 2, 2],
                "y": [1, 1, 2, 2, 1, 1, 2, 2],
                "a": [1, 2, 3, 4, 5, 6, 7, 8],
                "b": [11, 12, 13, 14, 15, 16, 17, 18],
            },
        )
        return df.set_index(["x", "y"])


class MultiColumn(FrameContainer):
    row: int = 2
    column: int = 11

    @classmethod
    def dataframe(cls) -> DataFrame:
        a = ["a"] * 8 + ["b"] * 8
        b = (["c"] * 4 + ["d"] * 4) * 2
        c = np.repeat(range(1, 9), 2)
        d = ["x", "y"] * 8
        df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
        df.columns = pd.MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
        return df


class MultiIndexColumn(FrameContainer):
    row: int = 13
    column: int = 21

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "x": [1, 1, 1, 1, 2, 2, 2, 2],
                "y": [1, 1, 2, 2, 1, 1, 2, 2],
                "z": [1, 1, 1, 2, 2, 1, 1, 1],
                "c": [1, 2, 3, 4, 5, 6, 7, 8],
                "d": [11, 12, 13, 14, 15, 16, 17, 18],
                "e": [21, 22, 23, 24, 25, 26, 27, 28],
                "f": [31, 32, 33, 34, 35, 36, 37, 38],
            },
        )
        df = df.set_index(["x", "y", "z"])
        x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
        df.columns = pd.MultiIndex.from_tuples(x, names=["a", "b"])
        return df


class WideColumn(FrameContainer):
    row: int = 3
    column: int = 29

    @classmethod
    def dataframe(cls) -> DataFrame:
        x = ["i", "i", "j", "j", "i"]
        y = ["k", "l", "k", "l", "k"]
        a = list(range(5))
        b = list(range(10, 15))
        df = DataFrame({"x": x, "y": y, "a": a, "b": b})
        return df.set_index(["x", "y"])

    def init(self) -> None:
        self.sf.add_wide_column("u", list(range(3)))
        self.sf.add_wide_column("v", list(range(4)), style=True)


if __name__ == "__main__":
    sheet = create_sheet()

    classes: list[type[FrameContainer]] = [
        NoIndex,
        Index,
        MultiIndex,
        MultiColumn,
        MultiIndexColumn,
        WideColumn,
    ]

    fcs = [cls(sheet, style=True) for cls in classes]
    for fc in fcs:
        fc.sf.set_adjacent_column_width(1)

    fcs[3].sf.number_format("0.0", s="a", r=3)

    sf = fcs[-1].sf
    sf.autofit()

    fc = MultiIndex(sheet, row=15, style=True)
    sf = fc.sf
    sf.add_column("c", 10)
    sf.add_column("d", list(range(8)))
    sf.add_formula_column("e", "={c}+{d}")
    sf.style().autofit()
    sf.add_wide_column("u", list(range(3)))
    sf.add_wide_column("v", list(range(4)), number_format="0.0")
    sf.add_formula_column("u", "={u}+{a}")
    sf.add_formula_column("v", "={v}+{b}", style=True, autofit=True)
    sf.number_format(autofit=True, a="0.0", x="0.00")
