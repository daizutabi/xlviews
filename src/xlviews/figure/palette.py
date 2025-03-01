from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Hashable
from itertools import cycle, islice
from typing import TYPE_CHECKING, TypeAlias

from xlviews.chart.style import COLORS, MARKER_DICT

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from pandas import DataFrame


def get_columns_default(
    data: DataFrame,
    columns: str | list[str],
    default: dict[Hashable, str] | list[str] | None = None,
) -> tuple[list[str], dict[Hashable, str]]:
    if isinstance(columns, str):
        columns = [columns]

    if any(c not in data for c in columns):
        data = data.drop_duplicates()
        values = [tuple(t) for t in data.itertuples(index=False)]
        default = dict(zip(values, cycle(columns), strict=False))
        return data.columns.tolist(), default

    if default is None:
        return columns, {}

    if isinstance(default, dict):
        return columns, default

    if isinstance(default, str):
        default = [default]

    data = data[columns].drop_duplicates()
    values = [tuple(t) for t in data.itertuples(index=False)]
    default = dict(zip(values, cycle(default), strict=False))

    return columns, default


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


class Palette(ABC):
    """A palette of items."""

    columns: list[str]
    index: dict[tuple[Hashable, ...], int]
    items: list[str]

    def __init__(
        self,
        data: DataFrame,
        columns: str | list[str],
        default: dict[Hashable, str] | list[str] | None = None,
    ) -> None:
        self.columns, default = get_columns_default(data, columns, default)
        self.index = get_index(data[self.columns], default)
        defaults = default.values()

        n = len(self.index) - len(default)
        self.items = [*defaults, *islice(self.cycle(defaults), n)]

    @abstractmethod
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:
        """Generate an infinite iterator of items."""

    def get(self, value: Hashable) -> int:
        if not isinstance(value, tuple):
            value = (value,)

        return self.index[value]

    def __getitem__(self, value: Hashable | dict) -> str:
        if value == {None: 0}:  # from series
            return self.items[0]

        if isinstance(value, dict):
            value = tuple(value[k] for k in self.columns)

        return self.items[self.get(value)]


class MarkerPalette(Palette):
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


class ColorPalette(Palette):
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


PaletteStyle: TypeAlias = (
    str
    | list[str]
    | dict[Hashable, str]
    | tuple[str | list[str], dict[Hashable, str] | list[str]]
    | Palette
)


def get_palette(
    cls: type[Palette],
    data: DataFrame,
    style: PaletteStyle | None,
) -> Palette | None:
    """Get a palette from a style."""
    if isinstance(style, Palette):
        return style

    if style is None:
        return None

    if isinstance(style, dict):
        return cls(data, data.columns.to_list(), style)

    if isinstance(style, tuple):
        return cls(data, *style)

    return cls(data, style)
