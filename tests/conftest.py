from __future__ import annotations

import pytest
import xlwings
from xlwings import App, Book

from xlviews.style import hide_gridlines


@pytest.fixture(scope="session")
def app():
    with xlwings.App(visible=False) as app:
        yield app


@pytest.fixture(scope="session")
def book(app: App):
    book = app.books.add()

    yield book

    book.close()


@pytest.fixture(scope="module")
def sheet_module(book: Book):
    sheet = book.sheets.add()
    hide_gridlines(sheet)

    yield sheet

    sheet.delete()


@pytest.fixture
def sheet(book: Book):
    sheet = book.sheets.add()

    yield sheet

    sheet.delete()
