from __future__ import annotations

from itertools import chain
from typing import TYPE_CHECKING

from .range import Range, iter_addresses

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence
    from typing import Self

    from xlwings import Range as RangeImpl
    from xlwings import Sheet


class RangeCollection:
    ranges: list[Range]

    def __init__(self, ranges: Iterable) -> None:
        self.ranges = list(chain.from_iterable(_iter_ranges(r) for r in ranges))

    def __repr__(self) -> str:
        cls = self.__class__.__name__
        addr = self.get_address(row_absolute=True, column_absolute=True)
        return f"<{cls} {addr}>"

    @classmethod
    def from_index(
        cls,
        row: int | Sequence[int | tuple[int, int]],
        column: int | Sequence[int | tuple[int, int]],
        sheet: Sheet | None = None,
    ) -> Self:
        return cls(_iter_ranges_from_index(row, column, sheet))

    def __len__(self) -> int:
        return len(self.ranges)

    def __iter__(self) -> Iterator[Range]:
        return iter(self.ranges)

    def get_address(
        self,
        *,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
    ) -> str:
        it = self.iter_addresses(
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
        )
        return ",".join(it)

    def iter_addresses(
        self,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        cellwise: bool = False,
        formula: bool = False,
    ) -> Iterator[str]:
        return iter_addresses(
            self,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            cellwise=cellwise,
            formula=formula,
        )

    @property
    def api(self):  # noqa: ANN201
        api = self.ranges[0].api

        if len(self.ranges) == 1:
            return api

        sheet = self.ranges[0].sheet
        union = sheet.book.app.api.Union

        for r in self.ranges[1:]:
            api = union(api, r.api)

        return api


def _iter_ranges(cell: str | Range | RangeImpl | tuple[int, int]) -> Iterator[Range]:
    if isinstance(cell, Range):
        yield cell

    elif isinstance(cell, str):
        for c in cell.split(","):
            yield Range(c)

    else:
        yield Range(cell)


def _iter_ranges_from_index(
    row: int | Sequence[int | tuple[int, int]],
    column: int | Sequence[int | tuple[int, int]],
    sheet: Sheet | None = None,
) -> Iterator[Range]:
    if isinstance(row, int) and isinstance(column, int):
        yield Range((row, column), sheet=sheet)

    elif isinstance(row, int) and not isinstance(column, int):
        for c in column:
            start, end = _unpack(c)
            yield Range((row, start), (row, end), sheet=sheet)

    elif isinstance(column, int) and not isinstance(row, int):
        for r in row:
            start, end = _unpack(r)
            yield Range((start, column), (end, column), sheet=sheet)

    else:
        msg = f"Either row or column must be an integer: {row=}, {column=}"
        raise TypeError(msg)


def _unpack(index: int | tuple[int, int]) -> tuple[int, int]:
    if isinstance(index, int):
        return index, index
    return index
