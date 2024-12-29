from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, overload

import xlwings as xw

if TYPE_CHECKING:
    from xlwings import App, Book, Sheet


@overload
def open_or_create(
    file: str | Path,
    app: App | None = None,
    sheet_name: None = None,
    *,
    visible: bool = True,
) -> Book: ...


@overload
def open_or_create(
    file: str | Path,
    app: App | None = None,
    sheet_name: str | None = None,
    *,
    visible: bool = True,
) -> Sheet: ...


def open_or_create(
    file: str | Path,
    app: App | None = None,
    sheet_name: str | None = None,
    *,
    visible: bool = True,
) -> Book | Sheet:
    """Open or create an Excel file.

    Args:
        path (str | Path): The path to the Excel file.
        app (App): The application to use.
        sheetname (str): The name of the sheet.
        visible (bool): Whether the file is visible.

    Returns:
        Book | Sheet: The book or sheet.
    """
    app = app or xw.apps.active or xw.apps.add()

    if Path(file).exists():
        book = app.books.open(file)
        created = False

    else:
        book = app.books.add()
        book.save(file)
        created = True

    app.visible = visible

    if sheet_name is None:
        return book

    if created:
        sheet: Sheet = book.sheets[0]
        sheet.name = sheet_name
        book.save()

        return sheet

    for sheet in book.sheets:
        if sheet.name == sheet_name:
            return sheet

    sheet = book.sheets.add(sheet_name, after=sheet)
    book.save()

    return sheet
