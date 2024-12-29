from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, overload

import xlwings as xw

if TYPE_CHECKING:
    from xlwings import App, Book, Sheet


def get_app() -> App:
    return xw.apps.active or xw.apps.add()


def get_book(name: str | None = None, app: App | None = None) -> Book:
    app = app or get_app()

    if not name:
        if app.books:
            return app.books.active

        return app.books.add()

    for book in app.books:
        if book.name == name:
            return book

    msg = f"Book {name!r} not found"
    raise ValueError(msg)


def get_sheet(
    name: str | None = None,
    book: Book | None = None,
    app: App | None = None,
) -> Sheet:
    book = book or get_book(app=app)

    if not name:
        return book.sheets.active

    for sheet in book.sheets:
        if sheet.name == name:
            return sheet

    return book.sheets.add(name, after=sheet)


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
    app = app or get_app()

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
        sheet = book.sheets[0]
        sheet.name = sheet_name
        book.save()

        return sheet

    sheet = get_sheet(sheet_name, book)
    book.save()

    return sheet
