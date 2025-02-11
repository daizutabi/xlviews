from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.testing.common import FrameContainer, create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer
from xlviews.testing.sheet_frame.pivot import create_base, create_multi

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class BaseParent(FrameContainer):
    @classmethod
    def dataframe(cls) -> DataFrame:
        return create_base()


class Base(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("v", "y", "x", formula=True)

    def init(self) -> None:
        super().init()
        self.sf.label = "v"


class MultiIndexParent(FrameContainer):
    column: int = 14

    @classmethod
    def dataframe(cls) -> DataFrame:
        return create_multi()


class MultiIndex(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("v", ["Y", "y"], ["X", "x"], formula=True)


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
