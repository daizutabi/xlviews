from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from xlwings import Range, Sheet

    from .range_collection import RangeCollection


if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator


def reference(cell: str | tuple[int, int] | Range, sheet: Sheet | None = None) -> str:
    """Return a reference to a cell with sheet name for chart."""
    if isinstance(cell, str):
        return cell

    if isinstance(cell, tuple):
        if sheet is None:
            raise ValueError("`sheet` is required when `cell` is a tuple")

        cell = sheet.range(*cell)

    return "=" + cell.get_address(include_sheetname=True)


def iter_address(
    ranges: Iterable[str | Range | RangeCollection],
    **kwargs,
) -> Iterator[str]:
    for rng in ranges:
        if isinstance(rng, str):
            yield rng
        else:
            yield rng.get_address(**kwargs)
