from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer
from xlviews.testing.sheet_frame.pivot import Base as BaseParent
from xlviews.testing.sheet_frame.pivot import MultiIndex as MultiIndexParent_

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class Base(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("v", "y", "x", formula=True)

    def init(self) -> None:
        cb = self.sf.colorbar(label="v", autofit=True)
        cb.set_adjacent_column_width(1)
        self.sf.style(color="#e0e0e0").autofit()


class MultiIndexParent(MultiIndexParent_):
    column: int = 15

    def init(self) -> None:
        self.sf.number_format("0.0").autofit()


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
