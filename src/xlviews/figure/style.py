from __future__ import annotations

from collections.abc import Callable, Hashable
from typing import TYPE_CHECKING, Any, TypeAlias, TypeVar

import pandas as pd
from pandas import DataFrame
from xlwings.constants import ChartType

from xlviews.dataframes.sheet_frame import SheetFrame

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from xlwings import Range, Sheet

    from xlviews.chart.series import Series
    from xlviews.core.range_collection import RangeCollection
    from xlviews.dataframes.groupby import GroupBy


Label: TypeAlias = str | dict[Hashable, str] | Callable[[Hashable], str] | None
Style: TypeAlias = (
    str | int | dict[Hashable, str | int] | Callable[[Hashable], str | int] | None
)


def format_label(label: Label, name: Hashable | None) -> str | None:
    if label is None:
        return None

    if isinstance(label, str):
        if name is None:
            return label
        if isinstance(name, tuple):
            return label.format(*name)
        return label.format(name)

    if isinstance(label, dict):
        return label[name]

    if callable(label):
        return label(name)

    msg = f"Invalid label: {label}"
    raise ValueError(msg)


def format_style(style: Style, name: Hashable | None) -> str | int | None:
    if style is None or isinstance(style, int | str):
        return style

    if isinstance(style, dict):
        return style[name]

    if callable(style):
        return style(name)

    msg = f"Invalid style: {style}"
    raise ValueError(msg)
