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


# class Style:

# Style: TypeAlias = (
#     float | str|list[str]|dict[Hashable, str | int] | tuple[tuple[str,...], dict[Hashable, str]]
# )

# def make_style(style: Style,names:list[str],index:list[tuple[Hashable,...]])->tuple[tuple[str,...], dict[Hashable, str]]:

Label: TypeAlias = str | Callable[[dict[str, Hashable]], str] | None


def format_label(label: Label, key: dict[str, Hashable]) -> str | None:
    if label is None:
        return None

    if isinstance(label, str):
        return label.format(**key)

    if callable(label):
        return label(key)

    msg = f"Invalid label: {label}"
    raise ValueError(msg)


# def format_style(style: Style, key: dict[str, Hashable]) -> str | int | None:
#     return None
#     if style is None or isinstance(style, int | str):
#         return style

#     if isinstance(style, dict):
#         return style[name]

#     if callable(style):
#         return style(name)

#     msg = f"Invalid style: {style}"
#     raise ValueError(msg)
