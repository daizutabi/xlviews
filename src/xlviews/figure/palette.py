from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Callable, Hashable
from itertools import cycle, islice
from typing import TYPE_CHECKING, Generic, TypeAlias, TypeVar

from pandas import MultiIndex

from xlviews.chart.style import COLORS, MARKER_DICT

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from pandas import DataFrame


T = TypeVar("T")


def get_columns_default(
    data: DataFrame,
    columns: str | list[str],
    default: dict[Hashable, T] | list[T] | None = None,
) -> tuple[list[str], dict[Hashable, T]]:
    if isinstance(columns, str):
        columns = [columns]

    if default is None:
        return columns, {}

    if isinstance(default, dict):
        return columns, default

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


class Palette(ABC, Generic[T]):
    """A palette of items."""

    columns: list[str]
    index: dict[tuple[Hashable, ...], int]
    items: list[T]

    def __init__(
        self,
        data: DataFrame,
        columns: str | list[str],
        default: dict[Hashable, T] | list[T] | None = None,
    ) -> None:
        self.columns, default = get_columns_default(data, columns, default)
        self.index = get_index(data[self.columns], default)
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

    def __getitem__(self, key: dict) -> T:
        if key == {None: 0}:  # from series
            return self.items[0]

        value = tuple(key[k] for k in self.columns)

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


class FunctionPalette(Generic[T]):
    columns: str | list[str]
    func: Callable[[Hashable], T]

    def __init__(self, columns: str | list[str], func: Callable[[Hashable], T]) -> None:
        self.columns = columns
        self.func = func

    def __getitem__(self, key: dict) -> T:
        if isinstance(self.columns, str):
            return self.func(key[self.columns])

        value = tuple(key[k] for k in self.columns)
        return self.func(value)


PaletteStyle: TypeAlias = (
    str
    | list[str]
    | dict[Hashable, str]
    | Callable[[Hashable], str]
    | tuple[str | list[str], list[str] | dict[Hashable, str]]
    | tuple[str | list[str], Callable[[Hashable], str]]
    | Palette
    | FunctionPalette
)


def get_palette(
    cls: type[Palette],
    data: DataFrame,
    style: PaletteStyle | None,
) -> Palette | FunctionPalette | None:
    """Get a palette from a style."""
    if isinstance(style, Palette | FunctionPalette):
        return style

    if style is None:
        return None

    if isinstance(style, Callable):
        if isinstance(data.index, MultiIndex):
            return FunctionPalette(data.index.names, style)
        return FunctionPalette(data.index.name, style)

    if data.index.name is not None or isinstance(data.index, MultiIndex):
        data = data.index.to_frame(index=False)

    if isinstance(style, dict):
        return cls(data, data.columns.to_list(), style)

    if isinstance(style, tuple):
        columns, default = style
        if callable(default):
            return FunctionPalette(columns, default)

        return cls(data, columns, default)

    columns = style

    if isinstance(columns, str):
        columns = [columns]

    if any(c not in data for c in columns):
        data = data.drop_duplicates()
        values = [tuple(t) for t in data.itertuples(index=False)]
        default = dict(zip(values, cycle(columns), strict=False))
        return cls(data, data.columns.tolist(), default)  # type: ignore

    return cls(data, columns)


def get_marker_palette(
    data: DataFrame,
    marker: PaletteStyle | None,
) -> MarkerPalette | FunctionPalette | None:
    return get_palette(MarkerPalette, data, marker)  # type: ignore


def get_color_palette(
    data: DataFrame,
    color: PaletteStyle | None,
) -> ColorPalette | FunctionPalette | None:
    return get_palette(ColorPalette, data, color)  # type: ignore
