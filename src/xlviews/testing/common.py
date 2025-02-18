from __future__ import annotations

from functools import cache
from typing import TYPE_CHECKING, Any

import xlwings
from pywintypes import com_error
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.style import hide_gridlines

if TYPE_CHECKING:
    from pandas import DataFrame
    from xlwings import Sheet


@cache
def is_app_available() -> bool:
    try:
        with xlwings.App(visible=False):
            pass
    except com_error:
        return False

    return True


def create_sheet() -> Sheet:
    for app in xlwings.apps:
        app.quit()

    book = xlwings.Book()
    sheet = book.sheets.add()
    sheet.range("A1").column_width = 1
    hide_gridlines(sheet)

    return sheet


def create_sheet_frame(
    df: DataFrame,
    sheet: Sheet,
    row: int,
    column: int,
    style: bool = False,
    table: bool = False,
) -> SheetFrame:
    sf = SheetFrame(row, column, df, sheet)
    if style:
        sf.style()
        sf.autofit()

    if table:
        sf.as_table()

    return sf


class FrameContainer:
    df: DataFrame
    sf: SheetFrame
    row: int = 2
    column: int = 2

    def __init__(
        self,
        sheet: Sheet,
        row: int = 0,
        column: int = 0,
        style: bool = False,
        **kwargs,
    ) -> None:
        self.df = self.dataframe()
        self.row = row or self.row
        self.column = column or self.column
        self.sf = create_sheet_frame(
            self.df,
            sheet,
            self.row,
            self.column,
            style=style,
            **self.kwargs(**kwargs),
        )
        self.init()

    def kwargs(self, **kwargs) -> dict[str, Any]:
        return kwargs

    def init(self) -> None:
        pass

    @classmethod
    def dataframe(cls) -> DataFrame:
        raise NotImplementedError
