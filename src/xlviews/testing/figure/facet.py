from __future__ import annotations

from typing import TYPE_CHECKING

from xlviews.chart.axes import Axes
from xlviews.figure.palette import get_color_palette, get_marker_palette
from xlviews.figure.plot import Plot
from xlviews.testing.common import create_sheet
from xlviews.testing.sheet_frame.pivot import Pivot

if TYPE_CHECKING:
    from collections.abc import Iterator
    from typing import Any

    from xlviews.dataframes.sheet_frame import SheetFrame


def facet(sf: SheetFrame) -> Iterator[tuple[dict[str, Any], Plot]]:
    sf.set_adjacent_column_width(1)

    ax = Axes(2, 11)
    data = sf.groupby(["A", "B", "X"]).agg(include_sheetname=True)
    cp = get_color_palette(data, "X")
    mp = get_marker_palette(data, "A")

    for key, plot in Plot.facet(ax, data, index="B", columns="A"):
        plot.add("u", "v").set(color=cp, marker=mp, label="{X}", alpha=0.8)
        plot.axes.title = "{A}_{B}".format(**key)
        yield key, plot


if __name__ == "__main__":
    sheet = create_sheet()
    fc = Pivot(sheet, style=True)
    list(facet(fc.sf))
