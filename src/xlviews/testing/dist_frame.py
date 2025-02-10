from __future__ import annotations

from pandas import DataFrame

from xlviews.testing.common import FrameContainer, create_sheet


class Parent(FrameContainer):
    row: int = 3
    column: int = 2

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = DataFrame(
            {
                "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2],
                "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
                "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
                "b": [10, 4, 9, 14, 5, 4, 6, 3, 4, 9, 12, 13, 9, 2],
            },
        )
        return df.set_index(["x", "y"])


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Parent(sheet, 3, 2, style=True)
    fc.sf.dist_frame(["a", "b"], by=["x", "y"])
