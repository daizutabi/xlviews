from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings as xw
from xlwings import Range

from xlviews.formula import const
from xlviews.style import set_alignment, set_font, set_table_style

if TYPE_CHECKING:
    from xlwings import Sheet
    from xlwings._xlwindows import COMRetryObjectWrapper


class Table:
    sheet: Sheet
    cell: Range
    api: COMRetryObjectWrapper

    def __init__(
        self,
        rng: Range | None = None,
        *,
        autofit: bool = False,
        const_header: bool = False,
        style: bool = False,
        sheet: Sheet | None = None,
        api: COMRetryObjectWrapper | None = None,
    ) -> None:
        if isinstance(rng, Range):
            self.cell = rng[0]

            self.api = rng.sheet.api.ListObjects.Add(
                xw.constants.ListObjectSourceType.xlSrcRange,
                rng.api,
                None,
                xw.constants.YesNoGuess.xlYes,
            )
        elif sheet and api:
            self.api = api
            self.cell = sheet.range(api.Range.Row, api.Range.Column)  # type: ignore
        else:
            raise ValueError("Either rng or sheet and api must be provided")

        self.sheet = self.cell.sheet

        if autofit:
            self.header.api.EntireColumn.AutoFit()

        if const_header:
            self.add_const_header()

        if style:
            set_table_style(self)

    @property
    def column(self) -> Range:
        start = self.cell.offset(1)
        end = start.expand("down")
        return self.cell.sheet.range(start, end)

    @property
    def header(self) -> Range:
        end = self.cell.expand("right")
        return self.cell.sheet.range(self.cell, end)

    @property
    def const_header(self) -> Range:
        return self.header.offset(-1)

    @property
    def columns(self) -> list[str]:
        names = self.header.value

        if not isinstance(names, list):
            raise NotImplementedError

        return [c or "" for c in names]

    def add_const_header(self, clear: bool = False) -> None:
        """Write the filtered element above the header.

        Args:
            clear (bool, optional): If True, clear the header.
        """
        if clear:
            self.const_header.value = None
        else:
            self.const_header.value = const(self.column, "=")
            set_font(self.const_header, size=8, italic=True, color="blue")
            set_alignment(self.const_header, "center")

    def auto_filter(self, *args, clear: bool = False, **field_criteria) -> None:
        """Filter the table by the given criteria.

        Args:
            *args: The criteria to filter by.
            **field_criteria: The criteria to filter by.

        The criteria is a dictionary with the column name as the key
        and the criteria as the value. The criteria can be one of the following:
           - list: specify the elements.
           - tuple: specify the range of values.
           - None: clear the filter.
           - other: specify the value.
        """
        if clear:
            clear_filter = {x: None for x in self.columns}
            clear_filter.update(field_criteria)
            field_criteria = clear_filter

        if args and isinstance(args[0], dict):
            field_criteria.update(args[0])
        else:
            for name, criteria in zip(args[::2], args[1::2], strict=True):
                field_criteria[name] = criteria

        autofilter = self.api.Range.AutoFilter  # type: ignore
        operator = xw.constants.AutoFilterOperator

        columns = self.columns

        for name, criteria in field_criteria.items():
            field = columns.index(name) + 1

            if isinstance(criteria, list):
                autofilter(
                    Field=field,
                    Criteria1=[str(x) for x in criteria],
                    Operator=operator.xlFilterValues,
                )

            elif isinstance(criteria, tuple):
                autofilter(
                    Field=field,
                    Criteria1=f">={criteria[0]}",
                    Operator=operator.xlAnd,
                    Criteria2=f"<={criteria[1]}",
                )

            elif criteria is None:
                autofilter(Field=field)

            else:
                autofilter(Field=field, Criteria1=f"{criteria}")

    def unlist(self) -> None:
        """Unlist the SheetFrame."""
        self.auto_filter(clear=True)
        self.api.Unlist()
        self.add_const_header(clear=True)
