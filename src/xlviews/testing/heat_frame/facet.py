from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndexParent

if TYPE_CHECKING:
    from pandas import DataFrame


class Facet(MultiIndexParent):
    column: int = 2

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = super().dataframe()

        for a, b in [(1, 3), (2, 2), (2, 4)]:
            ac = df.index.get_level_values("X") == a
            bc = df.index.get_level_values("x") == b
            df = df[~(ac & bc)]

        for a, b in [(2, 1), (2, 2), (3, 3), (4, 1), (4, 2), (4, 3)]:
            ac = df.index.get_level_values("Y") == a
            bc = df.index.get_level_values("y") == b
            df = df[~(ac & bc)]

        return df


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Facet(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    sf = HeatFrame(2, 8, fc.sf, "v", ["X", "x"], ["Y", "y"])
    sf.set_adjacent_column_width(1)

    HeatFrame.facet(2, 21, fc.sf, "v", x="x", y="y", col="X", row="Y")
