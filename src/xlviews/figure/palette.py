from __future__ import annotations

from abc import ABC, abstractmethod
from itertools import islice
from typing import TYPE_CHECKING, TypeVar

from xlviews.chart.style import cycle_colors, cycle_markers

if TYPE_CHECKING:
    from collections.abc import Hashable, Iterable, Iterator

    from pandas import DataFrame


def get_index(
    data: DataFrame,
    default: Iterable[Hashable] | None = None,
) -> dict[tuple[Hashable, ...], int]:
    data = data.drop_duplicates()
    values = [tuple(t) for t in data.itertuples(index=False)]

    if default is None:
        return dict(zip(values, range(len(data)), strict=True))

    index = {}
    current_index = 0

    for default_value in default:
        if isinstance(default_value, tuple):
            value = default_value
        elif isinstance(default_value, list):
            value = tuple(default_value)
        else:
            value = (default_value,)

        if value in values:
            index[value] = current_index
            current_index += 1

    for value in values:
        if value not in index:
            index[value] = current_index
            current_index += 1

    return index


T = TypeVar("T")


class Palette[T](ABC):
    """A palette of items."""

    index: dict[tuple[Hashable, ...], int]
    items: list[T]

    def __init__(
        self,
        data: DataFrame,
        columns: str | Iterable[str],
        default: dict[Hashable, T] | None = None,
    ) -> None:
        data = data.reset_index()

        if isinstance(columns, str):
            columns = [columns]
        else:
            columns = list(columns)

        if default is None:
            default = {}

        self.index = get_index(data[columns], default)
        defaults = default.values()

        n = len(self.index) - len(default)
        self.items = [*defaults, *islice(self.cycle(defaults), n)]

    @abstractmethod
    def cycle(self, defaults: Iterable[T]) -> Iterator[T]:
        """Generate an infinite iterator of items."""

    def get(self, value: Hashable) -> int:
        if not isinstance(value, tuple):
            value = (value,)

        return self.index[value]

    def __getitem__(self, value: Hashable) -> T:
        return self.items[self.get(value)]


class MarkerPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:
        """Generate an infinite iterator of markers."""
        return cycle_markers(defaults)


class ColorPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:
        """Generate an infinite iterator of colors."""
        return cycle_colors(defaults)
