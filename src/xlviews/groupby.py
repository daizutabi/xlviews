from __future__ import annotations

from typing import TYPE_CHECKING, TypeVar

import numpy as np
from pandas import DataFrame, Series

if TYPE_CHECKING:
    from collections.abc import Iterable, Sequence


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
