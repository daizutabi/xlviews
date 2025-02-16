from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.dataframes.heat_frame import Colorbar, facet
from xlviews.testing.common import create_sheet
from xlviews.testing.heat_frame.common import HeatFrameContainer
from xlviews.testing.sheet_frame.pivot import Pivot

if TYPE_CHECKING:
    from pandas import DataFrame

    from xlviews.dataframes.sheet_frame import SheetFrame


class FacetParent(Pivot):
    pass


class Facet(HeatFrameContainer):
    @classmethod
    def dataframe(cls, sf: SheetFrame) -> DataFrame:
        return sf.pivot_table("u", ["B", "Y"], ["A", "X"], "mean", formula=True)

    def init(self) -> None:
        super().init()


if __name__ == "__main__":
    sheet = create_sheet()

    fc = FacetParent(sheet, style=True)
    sf = fc.sf
    sf.set_adjacent_column_width(1)
    fc = Facet(sf)
    fc.sf.set_adjacent_column_width(1)
    fc.sf.range.impl.number_format = "0"
    fc.sf.autofit()

    df = sf.pivot_table("u", ["B", "Y"], ["A", "X"], "mean", formula=True)
    g = facet(2, 18, df, "B", "A")
    g[0, 0].autofit()
    g[0, 0].set_adjacent_column_width(1)
    g[0, 1].autofit()
    g[0, 1].set_adjacent_column_width(1)
    print(g)

    print(sf.agg(None, "u"))
    rng = sf.get_range("u")
    cell = g[0, -1].get_adjacent_cell()

    cb = Colorbar(cell.row + 1, cell.column, 6)
    cb.set(vmin=rng, vmax=rng).autofit()
    for x in g.flat:
        cb.apply(x.range)

    df = sf.pivot_table("u", ["B", "Y", "y"], ["A", "X", "x"], formula=True)
    g = facet(20, 18, df, "B", "A")
    g[0, 0].autofit()
    g[0, 0].set_adjacent_column_width(1)
    g[0, 1].autofit()
    g[0, 1].set_adjacent_column_width(1)
    print(g)

    print(sf.agg(None, "u"))
    rng = sf.get_range("u")
    cell = g[0, -1].get_adjacent_cell()

    cb = Colorbar(cell.row + 1, cell.column, 6)
    cb.set(vmin=rng, vmax=rng).autofit()
    for x in g.flat:
        cb.apply(x.range)
