from __future__ import annotations

from abc import ABC, abstractmethod
from itertools import cycle, islice
from typing import TYPE_CHECKING, Generic, TypeVar

from xlviews.chart.style import COLORS, MARKER_DICT

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


class Palette(Generic[T], ABC):
    """A palette of items."""

    columns: list[str]
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

        self.columns = columns

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

    def __getitem__(self, value: Hashable | dict[str, Hashable]) -> T:
        if isinstance(value, dict):
            value = tuple(value[k] for k in self.columns)

        return self.items[self.get(value)]


class MarkerPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:
        """Generate an infinite iterator of markers."""
        return cycle_markers(defaults)


def cycle_markers(skips: Iterable[str] | None = None) -> Iterator[str]:
    """Cycle through the markers."""
    if skips is None:
        skips = []

    markers = (m for m in MARKER_DICT if m != "")
    for marker in cycle(markers):
        if marker not in skips:
            yield marker


class ColorPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:
        """Generate an infinite iterator of colors."""
        return cycle_colors(defaults)


def cycle_colors(skips: Iterable[str] | None = None) -> Iterator[str]:
    """Cycle through the colors."""
    if skips is None:
        skips = []

    for color in cycle(COLORS):
        if color not in skips:
            yield color
