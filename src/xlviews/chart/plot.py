from __future__ import annotations

from itertools import chain
from typing import TYPE_CHECKING, Any, TypeAlias

from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.dataframes.sheet_frame import SheetFrame

from .axes import Axes

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy

Style: TypeAlias = str | list[str] | dict[str | tuple[str, ...], Any] | None
Label: TypeAlias = str | dict[str | tuple[str, ...], str] | None


def iter_by(df: DataFrame, style: Style) -> Iterator[str]:
    names = [*df.index.names, *df.columns]

    if isinstance(style, str) and style in names:
        yield style

    elif isinstance(style, list):
        yield from (s for s in style if s in names)

    elif isinstance(style, dict):
        for keys in style:
            for key in keys if isinstance(keys, tuple) else (keys,):
                if key in names:
                    yield key


def get_by(df: DataFrame, styles: Iterable[Style]) -> list[str]:
    it = (iter_by(df, style) for style in styles)
    return sorted(set(chain.from_iterable(it)))


def plot(
    ax: Axes,
    data: DataFrame,
    x: str,
    y: str,
    color: Style = None,
    label: Label = None,
) -> Series:
    by = get_by(data, [color])
    if not by:
        return ax.add_series(data[x], data[y], label=label)
    series = ax.add_series(data[x], data[y], label=label, by=by)
    return Series(series)


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
