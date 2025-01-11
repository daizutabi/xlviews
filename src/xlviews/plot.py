from __future__ import annotations

from xlwings import Range, Sheet

from xlviews.axes import Axes
from xlviews.series import Series
from xlviews.sheetframe import SheetFrame


def _plot(
    ax: Axes,
    data: SheetFrame,
    x: str | tuple,
    y: str | tuple | None = None,
    *,
    label: str | tuple[int, int] | Range = "",
    chart_type: int | None = None,
) -> Series:
    print(data.range(x))
    # if label is None:
    #     label = ax.sheet.range("A1")
    # s = ax.add_series(data[x], data[y], label=label)
    # s.set(marker="o", line="-", edge_weight=1.25, line_weight=1.25)
    # s.delete()
