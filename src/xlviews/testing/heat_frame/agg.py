from __future__ import annotations

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndex


class Agg(MultiIndex):
    column: int = 2

    def init(self) -> None:
        self.src = self.sf.groupby(["X", "Y"]).agg("mean", formula=True)


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Agg(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    sf = HeatFrame(2, 8, data=fc.src, x="X", y="Y", value="v", sheet=sheet)
    sf.set_adjacent_column_width(1)
