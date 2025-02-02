from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
import pandas as pd
from pandas import DataFrame

from xlviews.testing.common import FrameContainer, create_sheet

if TYPE_CHECKING:
    from xlwings import Sheet


class NoIndex(FrameContainer):
    row: int = 2
    column: int = 2

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
    column: int = 13

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "c": [1, 2, 3, 4, 5],
                "d": [11, 12, 13, 14, 15],
                "e": [21, 22, 23, 24, 25],
                "f": [31, 32, 33, 34, 35],
            },
        )
        x = [("a1", "b1"), ("a1", "b2"), ("a2", "b1"), ("a2", "b2")]
        df.columns = pd.MultiIndex.from_tuples(x, names=["a", "b"])
        return df


class MultiIndexColumn(FrameContainer):
    row: int = 10
    column: int = 11

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


class Curve(FrameContainer):
    row: int = 2
    column: int = 19

    @classmethod
    def dataframe(cls) -> DataFrame:
        a = ["a"] * 8 + ["b"] * 8
        b = (["c"] * 4 + ["d"] * 4) * 2
        c = np.repeat(range(1, 9), 2)
        d = ["x", "y"] * 8
        df = DataFrame(np.arange(16 * 6).reshape(16, 6).T)
        df.columns = pd.MultiIndex.from_arrays([a, b, c, d], names=["s", "t", "r", "i"])
        return df


class WideColumn(FrameContainer):
    row: int = 3
    column: int = 37

    def init(self) -> None:
        self.sf.add_wide_column("u", range(3), autofit=True, style=True)
        self.sf.add_wide_column("v", range(4), autofit=True, style=True)

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame({"x": ["i", "j"], "y": ["k", "l"], "a": [1, 2], "b": [3, 4]})
        return df.set_index(["x", "y"])


def create(sheet: Sheet) -> list[FrameContainer]:
    classes = [
        NoIndex,
        Index,
        MultiIndex,
        MultiColumn,
        MultiIndexColumn,
        Curve,
        WideColumn,
    ]
    return FrameContainer.from_classes(classes, sheet, style=True)


if __name__ == "__main__":
    sheet = create_sheet()
    create(sheet)
