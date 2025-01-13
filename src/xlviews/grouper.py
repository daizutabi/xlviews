from __future__ import annotations

from typing import TYPE_CHECKING, TypeVar, overload

import numpy as np
from pandas import DataFrame, Series

from xlviews.utils import iter_columns

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence

    from xlwings import Range

    from xlviews.range import RangeCollection
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


class Grouper:
    sf: SheetFrame
    by: list[str]
    grouped: dict[tuple, list[tuple[int, int]]]

    def __init__(self, sf: SheetFrame, by: str | list[str] | None = None) -> None:
        self.sf = sf
        self.by = list(iter_columns(sf, by)) if by else []
        self.grouped = sf.groupby(self.by)

    def __len__(self) -> int:
        return len(self.grouped)

    def keys(self) -> Iterator[tuple]:
        yield from self.grouped.keys()

    def values(self) -> Iterator[list[tuple[int, int]]]:
        yield from self.grouped.values()

    def items(self) -> Iterator[tuple[tuple, list[tuple[int, int]]]]:
        yield from self.grouped.items()

    def __iter__(self) -> Iterator[tuple]:
        yield from self.keys()

    def __getitem__(self, key: tuple) -> list[tuple[int, int]]:
        return self.grouped[key]

    def range(self, column: str, key: tuple) -> RangeCollection:
        return self.sf.range(column, self[key])

    def first_range(self, column: str, key: tuple) -> Range:
        return self.sf.range(column, self[key][0][0])

    @overload
    def ranges(self, column: str) -> Iterator[RangeCollection]: ...

    @overload
    def ranges(self, column: tuple) -> Iterator[Range]: ...

    @overload
    def ranges(self, column: None = None, **kwargs) -> Iterator[Iterator[Range]]: ...

    def ranges(
        self,
        column: str | tuple | None = None,
        **kwargs,
    ) -> Iterator[RangeCollection | Range | Iterator[Range]]:
        if isinstance(column, str):
            for key in self:
                yield self.range(column, key)
        elif column is None:
            for column in self:
                yield self.ranges(column, **kwargs)
        else:
            kwargs.update(dict(zip(self.by, column, strict=True)))
            yield from self.sf.ranges(**kwargs)

    def first_ranges(self, column: str) -> Iterator[Range]:
        for key in self:
            yield self.first_range(column, key)
