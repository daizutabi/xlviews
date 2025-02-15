from __future__ import annotations

from typing import TYPE_CHECKING, Any

import numpy as np
import pandas as pd

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from numpy.typing import NDArray
    from pandas import DataFrame
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

    def get_loc(self, key: str) -> tuple[int, int]:
        start = [k for k, _ in self.to_list()].index(key)
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

    def __init__(
        self,
        index: Axes,
        wide_index: WideIndex | dict[str, Any] | None = None,
    ) -> None:
        self.index = index if isinstance(index, pd.Index) else pd.Index(index)

        if wide_index is None:
            wide_index = WideIndex()
        elif not isinstance(wide_index, WideIndex):
            wide_index = WideIndex(wide_index)

        self.wide_index = wide_index

    def __len__(self) -> int:
        return len(self.index) + len(self.wide_index)

    @property
    def names(self) -> list[str]:
        return self.index.names

    @property
    def nlevels(self) -> int:
        """Get the number of levels in the index.

        Examples:
            >>> index = Index(["a"], {"b": [1, 2, 3]})
            >>> index.nlevels
            1

            >>> index = Index(pd.MultiIndex.from_tuples([("a", 1), ("a", 2)]))
            >>> index.nlevels
            2
        """
        return self.index.nlevels

    def to_list(self) -> list[Any]:
        """Convert the index to a list.

        Examples:
            >>> index = Index(["a"], {"b": [1, 2, 3]})
            >>> index.to_list()
            ['a', ('b', 1), ('b', 2), ('b', 3)]
        """
        return [*self.index.to_list(), *self.wide_index.to_list()]

    def __iter__(self) -> Iterator[str]:
        """Iterate over the index.

        Examples:
            >>> index = Index(["a"], {"b": [1, 2, 3]})
            >>> list(index)
            ['a', 'b']
        """
        yield from self.index.to_list()
        yield from self.wide_index

    def __contains__(self, key: Any) -> bool:
        """Check if the index contains a key.

        Examples:
            >>> index = Index(["a"], {"b": [1, 2, 3]})
            >>> "a" in index
            True
            >>> "b" in index
            True
            >>> "c" in index
            False
            >>> ("b", 1) in index
            False
        """
        if not isinstance(key, str):
            return False

        return any(key in k for k in self)

    def append(self, key: str, values: Iterable[Any] | None = None) -> None:
        """Append a key to the index.

        Args:
            key (str): The key to append.
            values (Iterable[Any], optional): The values to append.

        Examples:
            >>> index = Index(["a"])
            >>> index.append("b")
            >>> index.to_list()
            ['a', 'b']

            >>> index.append("c", [1, 2, 3])
            >>> index.to_list()
            ['a', 'b', ('c', 1), ('c', 2), ('c', 3)]
        """
        if values is None:
            self.index = self.index.append(pd.Index([key]))
        else:
            self.wide_index.append(key, values)

    def get_loc(self, key: str, offset: int = 0) -> int | tuple[int, int]:
        """Get the location of a key. The end index is inclusive.

        Examples:
            >>> index = Index(["a", "b"], {"c": [1, 2, 3], "d": [4, 5]})
            >>> index.get_loc("a")
            0
            >>> index.get_loc("b", offset=10)
            11
            >>> index.get_loc("c")
            (2, 4)
            >>> index.get_loc("d", offset=20)
            (25, 26)
        """
        if key not in self.index:
            offset = len(self.index) + offset
            loc = self.wide_index.get_loc(key)
            return loc[0] + offset, loc[1] + offset - 1

        loc = self.index.get_loc(key)
        if isinstance(loc, int):
            return loc + offset

        raise NotImplementedError

    def get_indexer(
        self,
        columns: dict[str, Any] | None = None,
        offset: int = 0,
        **kwargs,
    ) -> NDArray[np.intp]:
        if self.index.nlevels == 1:
            raise NotImplementedError

        if columns is not None:
            kwargs.update(columns)

        idx = [self.index.get_level_values(k) == v for k, v in kwargs.items()]

        return np.where(np.all(idx, axis=0))[0] + offset

    def to_frame(self, index: bool = True) -> DataFrame:
        return self.index.to_frame(index=index)
