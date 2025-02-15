from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer
from xlviews.testing.sheet_frame.pivot import Pivot
from xlviews.utils import iter_group_ranges

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class FacetParent(Pivot):
    pass


class Facet(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("u", ["B", "Y"], ["A", "X"], "mean", formula=True)


if __name__ == "__main__":
    sheet = create_sheet()

    fc = FacetParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    fc = Facet(sf)
    fc.sf.set_adjacent_column_width(1)

    df = sf.pivot_table("u", ["B", "Y"], ["A", "X"], "mean", formula=True)
    print(df)
    # HeatFrame(2, 9, df).autofit()

    a = df.loc[[1], [1]]
    a.index = a.index.droplevel(0)
    a.columns = a.columns.droplevel(0)
    HeatFrame(2, 20, a).autofit()
    # print(list(iter_group_ranges(df.index.get_level_values(0))))
    # print(a)
