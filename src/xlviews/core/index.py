from __future__ import annotations

from typing import TYPE_CHECKING, Any, overload

import pandas as pd

if TYPE_CHECKING:
    from collections.abc import Iterable

    from pandas._typing import Axes


class WideIndex(dict[str, list[Any]]):
    """Represent a wide index."""

    def __len__(self) -> int:
        return sum(len(values) for values in self.values())

    @property
    def names(self) -> list[str]:
        return list(self.keys())

    def to_list(self) -> list[tuple[str, Any]]:
        return [(key, value) for key in self for value in self[key]]

    @overload
    def get_loc(self, key: tuple[str, Any]) -> int: ...

    @overload
    def get_loc(self, key: str) -> tuple[int, int]: ...

    def get_loc(self, key: str | tuple[str, Any]) -> int | tuple[int, int]:
        lst = self.to_list()

        if not isinstance(key, str):
            return lst.index(key)

        start = [k for k, _ in lst].index(key)
        stop = start + len(self[key])
        return start, stop

    def append(self, key: str, values: Iterable[Any]) -> None:
        if key in self:
            msg = f"key {key!r} already exists"
            raise ValueError(msg)

        self[key] = list(values)


class Index:
    index: pd.Index
    wide_index: WideIndex

    def __init__(self, index: Axes, wide_index: WideIndex | None = None) -> None:
        self.index = index if isinstance(index, pd.Index) else pd.Index(index)
        self.wide_index = wide_index or WideIndex()

    def __len__(self) -> int:
        return len(self.index) + len(self.wide_index)

    @property
    def names(self) -> list[str]:
        return self.index.names

    @property
    def nlevels(self) -> int:
        return self.index.nlevels

    def to_list(self) -> list[Any]:
        return [*self.index.to_list(), *self.wide_index.to_list()]

    def append(self, key: str, values: Iterable[Any]) -> None:
        if self.wide_index:
            self.wide_index.append(key, values)
        else:
            self.index = self.index.append(values)

    # def get_loc(self, key: str | tuple) -> int | tuple[int, int]:
    #     if key not in self.index:
    #         return self.wide_index.get_loc(key)

    #     return self.index.get_loc(key)
