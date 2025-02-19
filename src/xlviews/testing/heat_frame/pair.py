from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.core.range import Range
from xlviews.dataframes.heat_frame import Colorbar, HeatFrame
from xlviews.testing.common import create_sheet
from xlviews.testing.sheet_frame.pivot import Pivot

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterator
    from typing import Any

    from xlviews.dataframes.sheet_frame import SheetFrame


def pair(sf: SheetFrame) -> Iterator[tuple[dict[Hashable, Any], HeatFrame]]:
    sf.set_adjacent_column_width(1)

    rng = sf.get_range("u")
    cbu = Colorbar(3, 11, 12)
    cbu.set(vmin=rng, vmax=rng).autofit()
    rng = sf.get_range("v")
    cbv = Colorbar(16, 11, 12)
    cbv.set(vmin=rng, vmax=rng).autofit()
    cbv.set_adjacent_column_width(1)

    df = sf.pivot_table(["u", "v"], ["B", "Y", "y"], ["A", "X", "x"], formula=True)

    for key, frame in HeatFrame.pair(2, 13, df, index="B", columns="A", axis=1):
        frame.autofit()
        frame.set_adjacent_column_width(1)
        if key["value"] == "u":
            cbu.apply(frame.range)
        else:
            cbv.apply(frame.range)
        cell = Range(frame.row - 1, frame.column + 1)
        cell.value = str(key)
        yield key, frame


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Pivot(sheet, style=True)
    list(pair(fc.sf))
