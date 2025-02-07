from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range

from .range_collection import RangeCollection

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from xlwings import Sheet


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


def iter_addresses(
    ranges: Range | RangeCollection | Iterable[Range | RangeCollection],
    *,
    row_absolute: bool = True,
    column_absolute: bool = True,
    include_sheetname: bool = False,
    external: bool = False,
    formula: bool = False,
    cellwise: bool = False,
) -> Iterator[str]:
    if isinstance(ranges, Range | RangeCollection):
        ranges = [ranges]

    for rng in ranges:
        yield rng.get_address(
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
        )
