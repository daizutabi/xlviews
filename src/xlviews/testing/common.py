from __future__ import annotations

from functools import cache
from typing import TYPE_CHECKING

import xlwings as xw
from pywintypes import com_error
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.range.style import hide_gridlines

if TYPE_CHECKING:
    from pandas import DataFrame
    from xlwings import Sheet


@cache
def is_excel_installed() -> bool:
    try:
        with xw.App(visible=False):
            pass
    except com_error:
        return False

    return True


def create_sheet() -> Sheet:
    for app in xw.apps:
        app.quit()

    book = xw.Book()
    sheet = book.sheets.add()
    hide_gridlines(sheet)

    return sheet


def create_sheet_frame(
    df: DataFrame,
    sheet: Sheet,
    row: int,
    column: int,
    **kwargs,
) -> SheetFrame:
    return SheetFrame(row, column, data=df, sheet=sheet, **kwargs)
