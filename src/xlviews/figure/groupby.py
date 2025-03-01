from __future__ import annotations

from itertools import chain
from typing import TYPE_CHECKING, Any, TypeAlias

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from pandas import DataFrame


Style: TypeAlias = str | list[str] | dict[str | tuple[str, ...], Any] | None
Label: TypeAlias = str | dict[str | tuple[str, ...], str] | None


def iter_by(df: DataFrame, style: Style) -> Iterator[str]:
    names = [*df.index.names, *df.columns]

    if isinstance(style, str) and style in names:
        yield style

    elif isinstance(style, list):
        yield from (s for s in style if s in names)

    elif isinstance(style, dict):
        for keys in style:
            for key in keys if isinstance(keys, tuple) else (keys,):
                if key in names:
                    yield key


def get_by(df: DataFrame, styles: Iterable[Style]) -> list[str]:
    it = (iter_by(df, style) for style in styles)
    return sorted(set(chain.from_iterable(it)))
