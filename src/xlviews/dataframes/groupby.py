from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING, TypeVar

import numpy as np
from pandas import DataFrame, Index, MultiIndex, Series
from xlwings import Range as RangeImpl

from xlviews.core.formula import Func, aggregate
from xlviews.core.range import Range
from xlviews.core.range_collection import RangeCollection
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
    if isinstance(a, DataFrame):
        df = a.reset_index(drop=True)
    else:
        df = DataFrame(a).reset_index(drop=True)

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
    if sf.columns.nlevels != 1:
        raise NotImplementedError

    if not by:
        start = sf.row + sf.columns.nlevels
        end = start + len(sf) - 1
        return {(): [(start, end)]}

    if isinstance(by, list) or ":" in by:
        by = list(iter_columns(sf.index.names, by))

    values = sf.index.to_frame()[by]
    index = create_group_index(values, sort=sort)

    offset = sf.row + sf.columns.nlevels
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
        self.by = list(iter_columns(sf.index.names, by)) if by else []
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
        return self.keys()

    def __getitem__(self, key: tuple) -> list[tuple[int, int]]:
        return self.group[key]

    def index(
        self,
        *,
        as_address: bool = False,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Index:
        if not as_address:
            values = self.keys()
            df = DataFrame(values, columns=self.by)

        else:
            cs = self.sf.index.names
            column = self.sf.column
            idx = [cs.index(c) + column for c in self.by]

            agg = partial(
                self._agg,
                "first",
                row_absolute=row_absolute,
                column_absolute=column_absolute,
                include_sheetname=include_sheetname,
                external=external,
                formula=formula,
            )

            values = {c: agg(i) for c, i in zip(self.by, idx, strict=True)}
            df = DataFrame(values)

        if len(self.by) == 1:
            return Index(df.iloc[:, 0], name=self.by[0])

        return MultiIndex.from_frame(df)

    def agg(
        self,
        func: Func | dict | Sequence[Func] = None,
        columns: str | list[str] | None = None,
        as_address: bool = False,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> DataFrame:
        if self.sf.columns.nlevels != 1:
            raise NotImplementedError

        if isinstance(func, dict):
            columns = list(func.keys())
        elif isinstance(columns, str):
            columns = [columns]

        idx = self.sf.get_indexer(columns)

        if columns is None:
            columns = self.sf.columns.to_list()

        index = self.index(
            as_address=as_address,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        agg = partial(
            self._agg,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )

        if isinstance(func, dict):
            it = zip(func.values(), idx, strict=True)
            values = np.array([list(agg(f, i)) for f, i in it]).T
            return DataFrame(values, index=index, columns=columns)

        if func is None or isinstance(func, str | Range | RangeImpl):
            values = np.array([list(agg(func, i)) for i in idx]).T
            return DataFrame(values, index=index, columns=columns)

        values = np.array([list(agg(f, i)) for i in idx for f in func]).T
        columns_ = MultiIndex.from_tuples([(c, f) for c in columns for f in func])
        return DataFrame(values, index=index, columns=columns_)

    def _agg(self, func: Func, column: int, **kwargs) -> Iterator[str]:
        if func == "first":
            func = None
            for row in self.values():
                rng = Range((row[0][0], column), sheet=self.sf.sheet)
                yield aggregate(func, rng, **kwargs)
        else:
            for row in self.values():
                rng = RangeCollection(row, column, self.sf.sheet)
                yield aggregate(func, rng, **kwargs)
