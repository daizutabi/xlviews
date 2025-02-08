from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings
from xlwings import Range as RangeImpl

from .address import get_index, index_to_column_name, split_book, split_sheet

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator
    from typing import Self

    from xlwings import Sheet


class Range:
    row: int
    column: int
    row_end: int
    column_end: int
    sheet: Sheet

    def __init__(
        self,
        cell1: str | Range | RangeImpl | tuple[int, int] | None = None,
        cell2: str | Range | RangeImpl | tuple[int, int] | None = None,
        sheet: Sheet | None = None,
    ) -> None:
        if cell1 is None:
            return

        t1 = to_tuple(cell1, sheet)
        self.sheet, self.row, self.column = t1[:3]

        if cell2 is None:
            self.row_end, self.column_end = t1[3:]
        else:
            t2 = to_tuple(cell2, sheet)

            if t1[0] != t2[0]:
                msg = f"Cells are not in the same sheet: {t1[0]} != {t2[0]}"
                raise ValueError(msg)

            self.row_end, self.column_end = t2[3:]

        self.row = min(self.row, self.row_end)
        self.column = min(self.column, self.column_end)
        self.row_end = max(self.row, self.row_end)
        self.column_end = max(self.column, self.column_end)

    @classmethod
    def from_index(
        cls,
        start: tuple[int, int],
        end: tuple[int, int] | None = None,
        sheet: Sheet | None = None,
    ) -> Self:
        rng = cls()
        rng.sheet = sheet or xlwings.sheets.active
        rng.row, rng.column = start
        rng.row_end, rng.column_end = end or start
        return rng

    def __len__(self) -> int:
        return (self.row_end - self.row + 1) * (self.column_end - self.column + 1)

    def __iter__(self) -> Iterator[Self]:
        for row in range(self.row, self.row_end + 1):
            for column in range(self.column, self.column_end + 1):
                yield self.__class__.from_index((row, column), sheet=self.sheet)

    def __getitem__(self, key: int) -> Self:
        if key < 0:
            key += len(self)

        if key < 0 or key >= len(self):
            raise IndexError("Index out of range")

        row = self.row + key // (self.column_end - self.column + 1)
        column = self.column + key % (self.column_end - self.column + 1)
        return self.__class__.from_index((row, column), sheet=self.sheet)

    def __repr__(self) -> str:
        addr = self.get_address(include_sheetname=True, external=True)
        return f"<{self.__class__.__name__} {addr}>"

    def offset(self, row_offset: int = 0, column_offset: int = 0) -> Self:
        return self.__class__.from_index(
            (self.row + row_offset, self.column + column_offset),
            (self.row_end + row_offset, self.column_end + column_offset),
            sheet=self.sheet,
        )

    def get_address(
        self,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> str:
        it = iter_addresses(
            self,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            formula=formula,
        )
        return next(it)

    def iter_addresses(
        self,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
    ) -> Iterator[str]:
        return iter_addresses(
            self,
            row_absolute=row_absolute,
            column_absolute=column_absolute,
            include_sheetname=include_sheetname,
            external=external,
            cellwise=True,
            formula=formula,
        )

    @property
    def impl(self) -> RangeImpl:
        cell1 = (self.row, self.column)
        cell2 = (self.row_end, self.column_end)
        return self.sheet.range(cell1, cell2)

    @property
    def api(self):  # noqa: ANN201
        return self.impl.api


def to_tuple(
    cell: str | tuple[int, int] | Range | RangeImpl,
    sheet: Sheet | None = None,
) -> tuple[Sheet, int, int, int, int]:
    if isinstance(cell, Range):
        return cell.sheet, cell.row, cell.column, cell.row_end, cell.column_end

    if isinstance(cell, RangeImpl):
        row, column = cell.row, cell.column
        row_end, column_end = cell.last_cell.row, cell.last_cell.column
        return cell.sheet, row, column, row_end, column_end

    sheet = sheet or xlwings.sheets.active

    if isinstance(cell, tuple):
        row, column = cell
        return sheet, row, column, row, column

    if isinstance(cell, str):
        book_name, address = split_book(cell, sheet)
        sheet_name, address = split_sheet(address, sheet)

        if book_name != sheet.book.name:
            msg = f"Book name does not match: {book_name} != {sheet.book.name}"
            raise ValueError(msg)

        if sheet_name != sheet.name:
            if sheet_name not in [s.name for s in sheet.book.sheets]:
                msg = f"Sheet not found: {sheet_name}"
                raise ValueError(msg)

            sheet = sheet.book.sheets[sheet_name]

        index = get_index(address)
        return sheet, *index  # type: ignore

    msg = f"Invalid type: {type(cell)}"
    raise TypeError(msg)


def iter_addresses(
    ranges: Range | Iterable[Range],
    *,
    row_absolute: bool = True,
    column_absolute: bool = True,
    include_sheetname: bool = False,
    external: bool = False,
    cellwise: bool = False,
    formula: bool = False,
) -> Iterator[str]:
    if isinstance(ranges, Range):
        ranges = [ranges]

    for rng in ranges:
        for addr in _iter_addresses(
            rng,
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
    rng: Range,
    *,
    row_absolute: bool = True,
    column_absolute: bool = True,
    include_sheetname: bool = False,
    external: bool = False,
    cellwise: bool = False,
) -> Iterator[str]:
    rp = "$" if row_absolute else ""
    cp = "$" if column_absolute else ""

    if external:
        prefix = f"[{rng.sheet.book.name}]{rng.sheet.name}!"
    elif include_sheetname:
        prefix = f"{rng.sheet.name}!"
    else:
        prefix = ""

    if cellwise:
        columns = range(rng.column, rng.column_end + 1)
        cnames = [index_to_column_name(c) for c in columns]

        for row in range(rng.row, rng.row_end + 1):
            for column in cnames:
                yield f"{prefix}{cp}{column}{rp}{row}"

    elif rng.row == rng.row_end and rng.column == rng.column_end:
        yield f"{prefix}{cp}{index_to_column_name(rng.column)}{rp}{rng.row}"

    else:
        start = f"{cp}{index_to_column_name(rng.column)}{rp}{rng.row}"
        end = f"{cp}{index_to_column_name(rng.column_end)}{rp}{rng.row_end}"
        yield f"{prefix}{start}:{end}"
