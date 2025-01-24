from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator, Sequence
    from typing import Self

    from xlwings import Sheet


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
        return ",".join(
            rng.get_address(
                row_absolute=row_absolute,
                column_absolute=column_absolute,
                include_sheetname=include_sheetname,
                external=external,
            )
            for rng in self
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
