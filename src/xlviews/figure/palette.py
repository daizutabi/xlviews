from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Callable, Hashable, Mapping
from itertools import cycle, islice
from typing import TYPE_CHECKING, Any, overload

from pandas import MultiIndex

from xlviews.chart.style import COLORS, MARKER_DICT

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from pandas import DataFrame


@overload
def get_columns_default(
    data: DataFrame,
    columns: str | list[str],
    default: None = None,
) -> tuple[list[str], dict[Any, Any]]: ...


@overload
def get_columns_default[T](
    data: DataFrame,
    columns: str | list[str],
    default: list[T],
) -> tuple[list[str], dict[Any, T]]: ...


@overload
def get_columns_default[T, K](
    data: DataFrame,
    columns: str | list[str],
    default: Mapping[K, T],
) -> tuple[list[str], dict[K, T]]: ...


def get_columns_default[T, K](
    data: DataFrame,
    columns: str | list[str],
    default: Mapping[K, T] | list[T] | None = None,
) -> tuple[list[str], dict[Any, Any]]:
    if isinstance(columns, str):
        columns = [columns]

    if default is None:
        return columns, {}

    if isinstance(default, Mapping):
        return columns, default if isinstance(default, dict) else dict(default)

    data = data[columns].drop_duplicates()
    values = [tuple(t) for t in data.itertuples(index=False)]
    default_ = dict(zip(values, cycle(default), strict=False))

    return columns, default_


def get_index(
    data: DataFrame,
    default: Iterable[Any] | None = None,
) -> dict[tuple[Any, ...], int]:
    data = data.drop_duplicates()
    values = [tuple(t) for t in data.itertuples(index=False)]

    if default is None:
        return dict(zip(values, range(len(data)), strict=True))

    index: dict[tuple[Any, ...], int] = {}
    current_index = 0

    for default_value in default:
        if isinstance(default_value, tuple):
            value = default_value  # pyright: ignore[reportUnknownVariableType]
        elif isinstance(default_value, list):
            value = tuple(default_value)  # pyright: ignore[reportUnknownArgumentType, reportUnknownVariableType]
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


class Palette[T](ABC):
    """A palette of items."""

    columns: list[str]
    index: dict[tuple[Any, ...], int]
    items: list[T]

    def __init__(
        self,
        data: DataFrame,
        columns: str | list[str],
        default: Mapping[Any, T] | list[T] | None = None,
    ) -> None:
        self.columns, default = get_columns_default(data, columns, default)
        self.index = get_index(data[self.columns], default)
        defaults = default.values()

        n = len(self.index) - len(default)
        self.items = [*defaults, *islice(self.cycle(defaults), n)]

    @abstractmethod
    def cycle(self, defaults: Iterable[T]) -> Iterator[T]:
        """Generate an infinite iterator of items."""

    def get(self, value: Any) -> int:
        if not isinstance(value, tuple):
            value = (value,)

        return self.index[value]

    def __getitem__(self, key: Mapping[Any, Any]) -> T:
        if key == {None: 0}:  # from series
            return self.items[0]

        value = tuple(key[k] for k in self.columns)

        return self.items[self.get(value)]


class MarkerPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:  # noqa: PLR6301
        """Generate an infinite iterator of markers."""
        return cycle_markers(defaults)


def cycle_markers(skips: Iterable[str] | None = None) -> Iterator[str]:
    """Cycle through the markers."""
    if skips is None:
        skips = []

    markers = (m for m in MARKER_DICT if m)
    for marker in cycle(markers):
        if marker not in skips:
            yield marker


class ColorPalette(Palette[str]):
    def cycle(self, defaults: Iterable[str]) -> Iterator[str]:  # noqa: PLR6301
        """Generate an infinite iterator of colors."""
        return cycle_colors(defaults)


def cycle_colors(skips: Iterable[str] | None = None) -> Iterator[str]:
    """Cycle through the colors."""
    if skips is None:
        skips = []

    for color in cycle(COLORS):
        if color not in skips:
            yield color


class FunctionPalette[T]:
    columns: Hashable | list[Hashable | None] | None
    func: Callable[[Any], T]

    def __init__(
        self,
        columns: Hashable | list[Hashable | None] | None,
        func: Callable[[Any], T],
    ) -> None:
        self.columns = columns
        self.func = func

    def __getitem__(self, key: Mapping[Any, Any]) -> T:
        if not isinstance(self.columns, list):
            return self.func(key[self.columns])

        value = tuple(key[k] for k in self.columns)
        return self.func(value)


type PaletteStyle[T] = (
    str
    | list[str]
    | dict[Hashable, T]
    | Callable[[Any], T]
    | tuple[Any, list[T] | dict[Hashable, T]]
    | tuple[Any, Callable[[Any], T]]
    | Palette[T]
    | FunctionPalette[T]
)


def get_palette[T](
    cls: type[Palette[T]],
    data: DataFrame,
    style: PaletteStyle[T] | None,
) -> Palette[T] | FunctionPalette[T] | None:
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
        return cls(data, data.columns.tolist(), default)  # pyright: ignore[reportArgumentType]

    return cls(data, columns)


def get_marker_palette(
    data: DataFrame,
    marker: PaletteStyle[str] | None,
) -> MarkerPalette | FunctionPalette[str] | None:
    return get_palette(MarkerPalette, data, marker)  # pyright: ignore[reportReturnType]


def get_color_palette(
    data: DataFrame,
    color: PaletteStyle[str] | None,
) -> ColorPalette | FunctionPalette[str] | None:
    return get_palette(ColorPalette, data, color)  # pyright: ignore[reportReturnType]
