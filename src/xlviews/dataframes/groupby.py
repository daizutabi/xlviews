from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING, TypeVar, overload

import numpy as np
from pandas import DataFrame, MultiIndex, Series

from xlviews.range.formula import aggregate
from xlviews.range.range import Range
from xlviews.range.range_collection import RangeCollection
from xlviews.utils import iter_columns

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence

    from .sheet_frame import SheetFrame

H = TypeVar("H")
T = TypeVar("T")


def to_dict(keys: Iterable[H], values: Iterable[T]) -> dict[H, list[T]]:
    result = {}

    for key, value in zip(keys, values, strict=True):
        result.setdefault(key, []).append(value)

    return result


def create_group_index(
    a: Sequence | Series | DataFrame,
    sort: bool = True,
) -> dict[tuple, list[tuple[int, int]]]:
    df = a.reset_index(drop=True) if isinstance(a, DataFrame) else DataFrame(a)

    dup = df[df.ne(df.shift()).any(axis=1)]

    start = dup.index.to_numpy()
    end = np.r_[start[1:] - 1, len(df) - 1]

    keys = [tuple(v) for v in dup.to_numpy()]
    values = [(int(s), int(e)) for s, e in zip(start, end, strict=True)]

    index = to_dict(keys, values)

    if not sort:
        return index

    return dict(sorted(index.items()))


def groupby(
    sf: SheetFrame,
    by: str | list[str] | None,
    *,
    sort: bool = True,
) -> dict[tuple, list[tuple[int, int]]]:
    """Group by the specified column and return the group key and row number."""
    if not by:
        if sf.columns_names is None:
            start = sf.row + sf.columns_level
            end = start + len(sf) - 1
            return {(): [(start, end)]}

        start = sf.column + 1
        end = start + len(sf.value_columns) - 1
        return {(): [(start, end)]}

    if sf.columns_names is None:
        if isinstance(by, list) or ":" in by:
            by = list(iter_columns(sf, by))
        values = sf.data.reset_index()[by]

    else:
        df = DataFrame(sf.value_columns, columns=sf.columns_names)
        values = df[by]

    index = create_group_index(values, sort=sort)

    if sf.columns_names is None:
        offset = sf.row + sf.columns_level  # vertical
    else:
        offset = sf.column + sf.index_level  # horizontal

    return {k: [(x + offset, y + offset) for x, y in v] for k, v in index.items()}


class GroupBy:
    sf: SheetFrame
    by: list[str]
    group: dict[tuple, list[tuple[int, int]]]

    def __init__(
        self,
        sf: SheetFrame,
        by: str | list[str] | None = None,
        *,
        sort: bool = True,
    ) -> None:
        self.sf = sf
        self.by = list(iter_columns(sf, by)) if by else []
        self.group = groupby(sf, self.by, sort=sort)

    def __len__(self) -> int:
        return len(self.group)

    def keys(self) -> Iterator[tuple]:
        yield from self.group.keys()

    def values(self) -> Iterator[list[tuple[int, int]]]:
        yield from self.group.values()

    def items(self) -> Iterator[tuple[tuple, list[tuple[int, int]]]]:
        yield from self.group.items()

    def __iter__(self) -> Iterator[tuple]:
        yield from self.keys()

    def __getitem__(self, key: tuple) -> list[tuple[int, int]]:
        return self.group[key]

    @overload
    def range(self, columns: str, key: tuple) -> RangeCollection: ...

    @overload
    def range(self, columns: list[str] | None, key: tuple) -> list[RangeCollection]: ...

    def range(
        self,
        columns: str | list[str] | None,
        key: tuple,
    ) -> RangeCollection | list[RangeCollection]:
        if isinstance(columns, str):
            return self.range([columns], key)[0]

        idx = self.sf.column_index(columns)
        row = self[key]

        return [RangeCollection.from_index(row, i, self.sf.sheet) for i in idx]

    @overload
    def first_range(self, columns: str, key: tuple) -> Range: ...

    @overload
    def first_range(self, columns: list[str] | None, key: tuple) -> list[Range]: ...

    def first_range(
        self,
        columns: str | list[str] | None,
        key: tuple,
    ) -> Range | list[Range]:
        if isinstance(columns, str):
            return self.first_range([columns], key)[0]

        idx = self.sf.column_index(columns)
        row = self[key][0][0]

        return [Range((row, i), sheet=self.sf.sheet) for i in idx]

    @overload
    def ranges(self, columns: str) -> Iterator[RangeCollection]: ...

    @overload
    def ranges(self, columns: list[str] | None) -> Iterator[list[RangeCollection]]: ...

    def ranges(
        self,
        columns: str | list[str] | None = None,
    ) -> Iterator[RangeCollection | list[RangeCollection]]:
        for key in self:
            yield self.range(columns, key)

    def first_ranges(self, column: str) -> Iterator[Range]:
        for key in self:
            yield self.first_range(column, key)

    def index(
        self,
        *,
        as_address: bool = False,
        **kwargs,
    ) -> DataFrame:
        if as_address:
            values = {c: self._agg_column("first", c, **kwargs) for c in self.by}
            return DataFrame(values)

        values = self.keys()
        return DataFrame(values, columns=self.by)

    def agg(
        self,
        func: str | Range | None | dict | Sequence[str | Range | None],
        columns: str | Sequence[str] | None = None,
        as_address: bool = False,
        formula: bool = False,
        **kwargs,
    ) -> DataFrame:
        agg = partial(self._agg_column, formula=formula, **kwargs)

        index_df = self.index(as_address=as_address, formula=formula, **kwargs)
        index = MultiIndex.from_frame(index_df)

        if isinstance(func, dict):
            return DataFrame({c: agg(f, c) for c, f in func.items()}, index=index)

        if columns is None:
            columns = self.sf.value_columns
        elif isinstance(columns, str):
            columns = [columns]

        if func is None or isinstance(func, str | Range):
            return DataFrame({c: agg(func, c) for c in columns}, index=index)

        values = {(c, f): agg(f, c) for c in columns for f in func}
        return DataFrame(values, index=index)

    def _agg_column(
        self,
        func: str | Range | None,
        column: str,
        **kwargs,
    ) -> list[str]:
        if func == "first":
            ranges = self.first_ranges(column)
            func = None
        else:
            ranges = self.ranges(column)

        return [aggregate(func, rng, **kwargs) for rng in ranges]
