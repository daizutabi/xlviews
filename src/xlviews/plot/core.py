from __future__ import annotations

from typing import TYPE_CHECKING, Any, TypeAlias

import pandas as pd
from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.chart.axes import Axes
from xlviews.dataframes.sheet_frame import SheetFrame

from .style import format_label

if TYPE_CHECKING:
    from collections.abc import Callable, Hashable, Iterable, Iterator

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy


def plot_series(
    ax: Axes,
    s: pd.Series,
    x: str,
    y: str,
    label: str | None = None,
    # label: str | dict[Hashable, str] | Callable[[Hashable], str] = None,
    # marker: Style = None,
    # color: Style = None,
    # alpha: float | None = None,
    # weight: float | None = None,
    # size: int | None = None,
) -> Series:
    label = format_label(label, s.name)

    return ax.add_series(s[x], s[y], label=label)


# def plot(
#     ax: Axes,
#     data: DataFrame | pd.Series,
#     x: str,
#     y: str,
#     label: Label = None,
#     marker: Style = None,
#     color: Style = None,
#     alpha: float | None = None,
#     weight: float | None = None,
#     size: int | None = None,
# ) -> Series:
#     if isinstance(data, pd.Series):
#         label = label if isinstance(label, str) else None
#         marker = marker if isinstance(marker, str) else None
#         color = color if isinstance(color, str) else None
#         s = ax.add_series(data[x], data[y], label=label).set(marker, color)

#     by = get_by(data, [color])

#     if not by:
#         label = label if isinstance(label, str) else None
#         return ax.add_series(data[x], data[y], label=label)

#     for key, s in data.iterrows():
#         print(key, s)

# series = ax.add_series(data[x], data[y], label=label, by=by)
# return Series(series)


# def get_range(
#     data: SheetFrame | GroupBy,
#     column: str,
#     key: str | tuple | None = None,
# ) -> Range | RangeCollection:
#     if isinstance(data, SheetFrame):
#         return data.range(column)

#     if isinstance(key, str):
#         key = (key,)

#     return data.range(column, key or ())


# def get_label(
#     data: SheetFrame | GroupBy,
#     column: str,
#     key: str | tuple | None = None,
# ) -> Range:
#     if isinstance(data, SheetFrame):
#         return data.first_range(column)

#     if isinstance(key, str):
#         key = (key,)

#     return data.first_range(column, key or ())


# def plot(
#     data: SheetFrame | GroupBy,
#     x: str,
#     y: str | None = None,
#     *,
#     key: str | tuple | None = None,
#     ax: Axes | None = None,
#     label: str | tuple[int, int] | Range = "",
#     chart_type: int | None = None,
#     sheet: Sheet | None = None,
# ) -> Series:
#     ct = ChartType.xlXYScatter if chart_type is None else chart_type
#     ax = ax or Axes(chart_type=ct)

#     xrng = get_range(data, x, key)
#     yrng = get_range(data, y, key) if y else None

#     if isinstance(label, str):
#         label = get_label(data, label, key)

#     return ax.add_series(xrng, yrng, label=label, chart_type=chart_type, sheet=sheet)
