from __future__ import annotations

from typing import TYPE_CHECKING, Any, TypeAlias

import pandas as pd
from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.dataframes.sheet_frame import SheetFrame

from ..chart.axes import Axes

if TYPE_CHECKING:
    from collections.abc import Callable, Hashable, Iterable, Iterator

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy


def format_label(
    label: str | Callable[[Hashable], str | None] | None,
    name: Hashable | None,
) -> str | None:
    if label is None:
        return None

    if isinstance(label, str):
        if name is None:
            return label
        if isinstance(name, tuple):
            return label.format(*name)
        return label.format(name)

    return label(name)
