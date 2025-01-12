from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence
    from typing import Self

    from xlwings import Sheet


def reference(cell: str | tuple[int, int] | Range, sheet: Sheet | None = None) -> str:
    """Return a reference to a cell with sheet name."""
    if isinstance(cell, str):
        return cell

    if sheet is None:
        if isinstance(cell, tuple):
            raise ValueError("sheet is required when `cell` is a tuple")

        sheet = cell.sheet

    return "=" + sheet.range(*cell).get_address(include_sheetname=True)


def iter_ranges(
    sheet: Sheet,
    row: int | Sequence[int | tuple[int, int]],
    column: int | Sequence[int | tuple[int, int]],
) -> Iterator[Range]:
    if isinstance(row, int) and isinstance(column, int):
        yield sheet.range(row, column)
        return

    if isinstance(row, int) and not isinstance(column, int):
        axis = 0
        index = column
    elif isinstance(column, int) and not isinstance(row, int):
        axis = 1
        index = row
    else:
        msg = "Either row or column must be an integer."
        raise TypeError(msg)

    def get_range(start_end: int | tuple[int, int]) -> Range:
        if isinstance(start_end, int):
            start = end = start_end
        else:
            start, end = start_end

        if axis == 0:
            return sheet.range((row, start), (row, end))

        return sheet.range((start, column), (end, column))

    yield from (get_range(i) for i in index)


def union_api(ranges: Iterable[Range]):  # noqa: ANN201
    ranges = list(ranges)

    api = ranges[0].api

    if len(ranges) == 1:
        return api

    sheet = ranges[0].sheet
    union = sheet.book.app.api.Union

    for r in ranges[1:]:
        api = union(api, r.api)

    return api


def union(ranges: Iterable[Range]) -> Range:
    ranges = list(ranges)

    if len(ranges) == 1:
        return ranges[0]

    sheet = ranges[0].sheet

    return sheet.range(union_api(ranges).Address)


def multirange(
    sheet: Sheet,
    row: int | Sequence[int | tuple[int, int]],
    column: int | Sequence[int | tuple[int, int]],
) -> Range:
    """Create a discontinuous range.

    Either row or column must be an integer.
    If the other is not an integer, it is treated as a list.
    If index is (int, int), it is a simple range.
    Otherwise, each element of index is an int or (int, int), and they are
    concatenated to create a discontinuous range.

    Args:
        sheet (Sheet): The sheet object.
        row (int, tuple, or list): The row number.
        column (int, tuple, or list): The column number.

    Returns:
        Range: The discontinuous range.
    """
    return union(iter_ranges(sheet, row, column))


class RangeCollection:
    ranges: list[Range]

    def __init__(self, ranges: Iterable) -> None:
        self.ranges = []

        for rng in ranges:
            self.ranges.append(rng if isinstance(rng, Range) else Range(rng))

    def __repr__(self) -> str:
        cls = self.__class__.__name__
        addr = self.get_address(row_absolute=True, column_absolute=True)
        return f"<{cls} {addr}>"

    @classmethod
    def from_index(
        cls,
        sheet: Sheet,
        row: int | Sequence[int | tuple[int, int]],
        column: int | Sequence[int | tuple[int, int]],
    ) -> Self:
        return cls(iter_ranges(sheet, row, column))

    def __len__(self) -> int:
        return sum(len(rng) for rng in self.ranges)

    def __iter__(self) -> Iterator[Range]:
        for rng in self.ranges:
            yield from rng

    def first(self) -> Range:
        return next(iter(self))

    def get_address(
        self,
        *,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
    ) -> str:
        return ",".join(
            rng.get_address(
                row_absolute=row_absolute,
                column_absolute=column_absolute,
                include_sheetname=include_sheetname,
                external=external,
            )
            for rng in self.ranges
        )

    @property
    def api(self):  # noqa: ANN201
        return union_api(self.ranges)
