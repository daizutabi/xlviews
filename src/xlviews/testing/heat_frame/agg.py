from __future__ import annotations

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndex
from xlviews.utils import add_validate_list


class Agg(MultiIndex):
    column: int = 2


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Agg(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    sf = HeatFrame(2, 8, data=fc.sf, x="X", y="Y", value="v", aggfunc="mean")
    sf.set_adjacent_column_width(1)

    rng = sf.sheet.range("$M$13")
    add_validate_list(rng, ["min", "max", "mean", "median", "soa"], "mean")
    sf.sheet.range("$M$13").value = "max"
    sf = HeatFrame(8, 8, data=fc.sf, x="X", y="Y", value="v", aggfunc=rng)
