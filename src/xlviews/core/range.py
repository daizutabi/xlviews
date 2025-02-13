from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np
import xlwings
from pandas import DataFrame
from xlwings import Range as RangeImpl

from .address import index_to_column_name

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator
    from typing import Any, Self

    from pandas._typing import Axes
    from xlwings import Sheet


class Range:
    row: int
    column: int
    row_end: int
    column_end: int
    sheet: Sheet

    def __init__(
        self,
        cell1: tuple[int, int] | int,
        cell2: tuple[int, int] | int | None = None,
        sheet: Sheet | None = None,
    ) -> None:
        self.sheet = sheet or xlwings.sheets.active

        if isinstance(cell1, tuple):
            if not isinstance(cell2, tuple) and cell2 is not None:
                msg = "cell2 must be a tuple or None"
                raise TypeError(msg)

            self.row, self.column = cell1
            self.row_end, self.column_end = cell2 or cell1

        elif isinstance(cell2, int):
            self.row = self.row_end = cell1
            self.column = self.column_end = cell2

        else:
            msg = "cell2 must be an integer"
            raise TypeError(msg)

    @classmethod
    def from_range(cls, rng: RangeImpl) -> Self:
        start = rng.row, rng.column
        end = rng.last_cell.row, rng.last_cell.column
        return cls(start, end, rng.sheet)

    def __len__(self) -> int:
        return (self.row_end - self.row + 1) * (self.column_end - self.column + 1)

    def __iter__(self) -> Iterator[Self]:
        for row in range(self.row, self.row_end + 1):
            for column in range(self.column, self.column_end + 1):
                yield self.__class__((row, column), sheet=self.sheet)

    def __getitem__(self, key: int) -> Self:
        if key < 0:
            key += len(self)

        if key < 0 or key >= len(self):
            raise IndexError("Index out of range")

        row = self.row + key // (self.column_end - self.column + 1)
        column = self.column + key % (self.column_end - self.column + 1)
        return self.__class__((row, column), sheet=self.sheet)

    @property
    def last_cell(self) -> Self:
        return self[-1]

    def __repr__(self) -> str:
        addr = self.get_address(include_sheetname=True, external=True)
        return f"<{self.__class__.__name__} {addr}>"

    def offset(self, row_offset: int = 0, column_offset: int = 0) -> Self:
        return self.__class__(
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
    def value(self) -> Any:
        return self.impl.value

    @value.setter
    def value(self, value: Any) -> None:
        self.impl.value = value

    @property
    def api(self):  # noqa: ANN201
        return self.impl.api

    @property
    def frame(self) -> FrameRange:
        return FrameRange(
            (self.row, self.column),
            (self.row_end, self.column_end),
            sheet=self.sheet,
        )


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
        prefix = f"{prefix}{cp}"
        cindex = range(rng.column, rng.column_end + 1)
        cnames = [index_to_column_name(c) for c in cindex]

        rows = [f"{rp}{row}" for row in range(rng.row, rng.row_end + 1)]
        for row in rows:
            for cn in cnames:
                yield f"{prefix}{cn}{row}"

    elif rng.row == rng.row_end and rng.column == rng.column_end:
        yield f"{prefix}{cp}{index_to_column_name(rng.column)}{rp}{rng.row}"

    else:
        start = f"{cp}{index_to_column_name(rng.column)}{rp}{rng.row}"
        end = f"{cp}{index_to_column_name(rng.column_end)}{rp}{rng.row_end}"
        yield f"{prefix}{start}:{end}"


class FrameRange(Range):
    def get_address(
        self,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
        formula: bool = False,
        index: Axes | None = None,
        columns: Axes | None = None,
    ) -> DataFrame:
        rp = "$" if row_absolute else ""
        cp = "$" if column_absolute else ""
        f = "=" if formula else ""

        if external:
            prefix = f"{f}[{self.sheet.book.name}]{self.sheet.name}!{cp}"
        elif include_sheetname:
            prefix = f"{f}{self.sheet.name}!{cp}"
        else:
            prefix = f"{f}{cp}"

        cindex = range(self.column, self.column_end + 1)
        cnames = [index_to_column_name(c) for c in cindex]

        rows = [f"{rp}{row}" for row in range(self.row, self.row_end + 1)]
        values = np.array([f"{prefix}{cn}{row}" for row in rows for cn in cnames])
        values = values.reshape(len(rows), len(cnames))

        return DataFrame(values, index=index, columns=columns)
