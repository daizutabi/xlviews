# from __future__ import annotations

# from typing import TYPE_CHECKING

# from xlwings.constants import ChartType

# from xlviews.dataframes.sheet_frame import SheetFrame

# from .axes import Axes

# if TYPE_CHECKING:
#     from xlwings import Range, Sheet

#     from xlviews.chart.series import Series
#     from xlviews.core.range_collection import RangeCollection
#     from xlviews.dataframes.groupby import GroupBy


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
