from __future__ import annotations

from typing import TYPE_CHECKING, TypeVar

import numpy as np
from pandas import DataFrame, Series

from xlviews.range import RangeCollection
from xlviews.utils import iter_columns

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence

    from xlwings import Range

    from xlviews.sheetframe import SheetFrame

H = TypeVar("H")
T = TypeVar("T")


def _to_dict(keys: Iterable[H], values: Iterable[T]) -> dict[H, list[T]]:
    result = {}

    for key, value in zip(keys, values, strict=True):
        result.setdefault(key, []).append(value)

    return result


def create_group_index(
    a: Sequence | Series | DataFrame,
) -> dict[tuple, list[tuple[int, int]]]:
    df = a.reset_index(drop=True) if isinstance(a, DataFrame) else DataFrame(a)

    dup = df[df.ne(df.shift()).any(axis=1)]

    start = dup.index.to_numpy()
    end = np.r_[start[1:] - 1, len(df) - 1]

    keys = [tuple(v) for v in dup.to_numpy()]
    values = [(int(s), int(e)) for s, e in zip(start, end, strict=True)]

    return _to_dict(keys, values)


class GroupedRange:
    sf: SheetFrame
    by: list[str]
    grouped: dict[tuple, list[tuple[int, int]]]

    def __init__(self, sf: SheetFrame, by: str | list[str] | None = None) -> None:
        self.sf = sf
        self.by = list(iter_columns(sf, by)) if by else []
        self.grouped = sf.groupby(self.by)

    def iter_ranges(self, column: str) -> Iterator[RangeCollection]:
        col = self.sf.index(column)
        if not isinstance(col, int):
            raise NotImplementedError

        sheet = self.sf.sheet

        for row in self.grouped.values():
            yield RangeCollection.from_index(sheet, row, col)

    def iter_first_ranges(self, column: str) -> Iterator[Range]:
        col = self.sf.index(column)
        if not isinstance(col, int):
            raise NotImplementedError

        sheet = self.sf.sheet

        for row in self.grouped.values():
            start = row[0][0]
            yield sheet.range(start, col)
