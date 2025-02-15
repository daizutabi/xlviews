from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.dataframes.heat_frame import HeatFrame

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class HeatFrameContainer:
    row: int = 0
    column: int = 0
    df: DataFrame
    sf: HeatFrame

    def __init__(self, sf: SheetFrame, row: int = 0, column: int = 0) -> None:
        self.row = row or self.row or sf.row
        self.column = column or self.column or (sf.column + sf.width + 1)
        self.df = self.dataframe(sf)
        self.sf = HeatFrame(self.row, self.column, data=self.df, sheet=sf.sheet)
        self.init()

    def init(self) -> None:
        self.sf.autofit()

    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        raise NotImplementedError
