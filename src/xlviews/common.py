from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, overload

import xlwings as xw

if TYPE_CHECKING:
    from xlwings import App, Book, Range, Sheet


def get_app() -> App:
    """Get the active app or create a new one."""
    return xw.apps.active or xw.apps.add()


def get_book(name: str | None = None, *, app: App | None = None) -> Book:
    """Get a book from the active app or a specific app."""
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
    *,
    book: Book | None = None,
    app: App | None = None,
) -> Sheet:
    """Get a sheet from the active book or a specific book."""
    book = book or get_book(app=app)

    if not name:
        return book.sheets.active

    for sheet in book.sheets:
        if sheet.name == name:
            return sheet

    return book.sheets.add(name, after=sheet)


def get_range(
    *args,
    sheet: Sheet | None = None,
    book: Book | None = None,
    app: App | None = None,
) -> Range:
    """Get a range from the active sheet or a specific sheet."""
    match len(args):
        case 0:
            sheet = sheet or get_sheet(book=book, app=app)
            return sheet.range("A1")

        case 1:
            sheet = sheet or get_sheet(book=book, app=app)
            if isinstance(args[0], str):
                return sheet.range(args[0])

            return sheet.range(*args[0])

        case 2:
            if isinstance(args[0], str):
                sheet = get_sheet(args[0], book=book, app=app)
                return get_range(args[1:], sheet=sheet)

            sheet = sheet or get_sheet(book=book, app=app)
            return sheet.range(*args)

        case 3:
            sheet = get_sheet(args[0], book=book, app=app)
            return get_range(args[1:], sheet=sheet)

        case _:
            msg = f"Invalid number of arguments: {len(args)}"
            raise ValueError(msg)


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

    sheet = get_sheet(sheet_name, book=book)
    book.save()

    return sheet


def delete_charts(sheet: Sheet | None = None) -> None:
    """Delete all charts in the sheet."""
    sheet = sheet or xw.sheets.active

    for chart in sheet.charts:
        chart.delete()


def quit_apps() -> None:
    for app in xw.apps:
        app.quit()
