from __future__ import annotations

import random

from pandas import DataFrame

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.base import MultiIndex


class Facet(MultiIndex):
    column: int = 2

    @classmethod
    def dataframe(cls) -> DataFrame:
        df = super().dataframe()
        return df.sample(frac=0.8, random_state=0)

    def init(self) -> None:
        self.src = self.sf.get_address(formula=True)


if __name__ == "__main__":
    sheet = create_sheet()

    fc = Facet(sheet, style=True)
    df = fc.src
    gr = df.groupby(["X", "Y"])
    print([x.shape for (x, _), x in gr])
    print(gr.get_group((1, 1)))

    # sf = fc.sf
    # sf.set_adjacent_column_width(1)
    # sf = HeatFrame(2, 8, data=fc.src, x="X", y="Y", value="v", sheet=sheet)
    # sf.set_adjacent_column_width(1)
