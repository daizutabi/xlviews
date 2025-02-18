from __future__ import annotations

from typing import ClassVar

from pandas import DataFrame

from xlviews.dataframes.dist_frame import DistFrame
from xlviews.testing.common import FrameContainer, create_sheet


class Parent(FrameContainer):
    row: int = 3
    column: int = 2
    index: ClassVar[list[str] | str] = ["x", "y"]

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2],
                "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
                "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
                "b": [10, 4, 9, 14, 5, 4, 6, 3, 4, 9, 12, None, 9, 2],
            },
        )
        return df.set_index(cls.index)


if __name__ == "__main__":
    sheet = create_sheet()
    Parent.index = ["x", "y"]
    fc = Parent(sheet, 3, 2, style=True)
    fc.sf.set_adjacent_column_width(1)
    fc.sf.number_format(b="0.0")
    DistFrame(fc.sf, ["a", "b"], by=["x", "y"])
