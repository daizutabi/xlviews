from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range

from xlviews.utils import int_to_column_name

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
    cellwise: bool = False,
    formula: bool = False,
) -> Iterator[str]:
    for addr in _iter_addresses(
        ranges,
        row_absolute=row_absolute,
        column_absolute=column_absolute,
        include_sheetname=include_sheetname,
        external=external,
        cellwise=cellwise,
    ):
        if formula:
            yield "=" + addr
        else:
            yield addr


def _iter_addresses(
    ranges: Range | RangeCollection | Iterable[Range | RangeCollection],
    *,
    cellwise: bool = False,
    **kwargs,
) -> Iterator[str]:
    if isinstance(ranges, Range | RangeCollection):
        ranges = [ranges]

    for rng in ranges:
        ranges = [rng] if isinstance(rng, Range) else rng

        for r in ranges:
            if cellwise:
                yield from _iter_addresses_from_range(r, **kwargs)
            else:
                yield r.get_address(**kwargs)


def _iter_addresses_from_range(
    rng: Range,
    *,
    row_absolute: bool = True,
    column_absolute: bool = True,
    include_sheetname: bool = False,
    external: bool = False,
) -> Iterator[str]:
    rp = "$" if row_absolute else ""
    cp = "$" if column_absolute else ""

    if external:
        prefix = f"[{rng.sheet.book.name}]{rng.sheet.name}!"
    elif include_sheetname:
        prefix = f"{rng.sheet.name}!"
    else:
        prefix = ""

    columns = [int_to_column_name(c) for c in range(rng.column, rng[-1].column + 1)]

    for row in range(rng.row, rng[-1].row + 1):
        for column in columns:
            yield f"{prefix}{cp}{column}{rp}{row}"
