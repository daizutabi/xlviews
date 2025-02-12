from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndexParent
from xlviews.utils import iter_group_ranges

if TYPE_CHECKING:
    from pandas import DataFrame


class FacetParent(MultiIndexParent):
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

    fc = FacetParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)

    df = sf.pivot_table("u", ["Y", "y"], ["X", "x"], formula=True)
    HeatFrame(2, 9, df).autofit()

    a = df.loc[[1], [1]]
    a.index = a.index.droplevel(0)
    a.columns = a.columns.droplevel(0)
    HeatFrame(2, 22, a).autofit()
    print(list(iter_group_ranges(df.index.get_level_values(0))))
    print(a)
