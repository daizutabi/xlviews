from __future__ import annotations

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndex
from xlviews.utils import add_validate_list


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

    rng = sf.sheet.range("$M$13")
    add_validate_list(rng, ["min", "max", "mean", "median", "soa"], "mean")
    sf.sheet.range("$M$13").value = "max"
    src = fc.sf.groupby(["X", "Y"]).agg("$M$13", formula=True)
    sf = HeatFrame(8, 8, data=src, x="X", y="Y", value="v", sheet=sheet)
