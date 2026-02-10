from __future__ import annotations

from typing import TYPE_CHECKING, Any

from xlviews.core.range import Range
from xlviews.dataframes.colorbar import Colorbar
from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.sheet_frame.pivot import Pivot

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator

    from xlviews.dataframes.sheet_frame import SheetFrame


def facet(sf: SheetFrame) -> Iterator[tuple[dict[Hashable, Any], HeatFrame]]:
    sf.set_adjacent_column_width(1)

    rng = sf.get_range("u")
    cb = Colorbar(3, 11, 12)
    cb.set(vmin=rng, vmax=rng).autofit()
    cb.set_adjacent_column_width(1)

    df = sf.pivot_table("u", ["B", "Y"], ["A", "X"], aggfunc="mean", formula=True)

    for key, frame in HeatFrame.facet(2, 13, df, index="B"):
        frame.autofit()
        frame.set_adjacent_column_width(1)
        cb.apply(frame.range)
        cell = Range(frame.row - 1, frame.column + 1)
        cell.value = str(key)
        yield key, frame


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Pivot(sheet, style=True)
    list(facet(fc.sf))
