from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.core.range import Range
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer
from xlviews.testing.sheet_frame.pivot import MultiIndex
from xlviews.utils import add_validate_list

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class AggParent(MultiIndex):
    column: int = 2


class AggStr(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("v", "Y", "X", "mean", formula=True)


class AggRange(HeatFrameContainer):
    row: int = 8

    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        func = Range((13, 14), sheet=sf.sheet)
        add_validate_list(func, ["min", "max", "mean", "median", "soa"], "mean")
        return sf.pivot_table("v", "Y", "X", func, formula=True)


if __name__ == "__main__":
    sheet = create_sheet()

    fc = AggParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    fc = AggStr(sf)
    fc = AggRange(sf)
    fc.sf.set_adjacent_column_width(1)
