from __future__ import annotations

from xlviews.core.range import Range
from xlviews.dataframes.colorbar import Colorbar
from xlviews.testing.common import create_sheet

if __name__ == "__main__":
    sheet = create_sheet()

    rng = Range((2, 2), (3, 3))
    rng.value = [[1, 2], [3, 4]]
    cb = Colorbar(2, 5, 6, sheet=sheet)
    cb.set(vmin=rng, vmax=rng, label="T", autofit=True)
    cb.set_adjacent_column_width(1)

    cb = Colorbar(2, 7, 10, orientation="horizontal")
    cb.set(vmin=rng, vmax=rng, label="T", autofit=True)
    cb.set_adjacent_column_width(1)
