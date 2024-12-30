from pathlib import Path

import pytest
import xlwings as xw
from xlwings import App, Book, Sheet

from xlviews.common import open_or_create


def test_app():
    from xlviews.common import get_app

    assert isinstance(get_app(), App)


def test_book():
    from xlviews.common import get_book

    book = get_book()

    assert isinstance(book, Book)
    assert xw.books.active.name == book.name
    assert get_book().name == book.name
    assert get_book(book.name).name == book.name


def test_book_error():
    from xlviews.common import get_book

    with pytest.raises(ValueError, match="Book 'invalid' not found"):
        get_book("invalid")


def test_sheet():
    from xlviews.common import get_sheet

    sheet = get_sheet()

    assert isinstance(sheet, Sheet)
    assert get_sheet().name == sheet.name
    assert get_sheet(sheet.name).name == sheet.name
    assert get_sheet("New").name == "New"


def test_range(book: Book):
    from xlviews.common import get_range

    cells = get_range(book=book)

    assert cells.row == 1
    assert cells.column == 1
    assert cells.shape == (1, 1)


def test_range_str(book: Book):
    from xlviews.common import get_range

    cells = get_range("B4:D10", book=book)

    assert cells.row == 4
    assert cells.column == 2
    assert cells.shape == (7, 3)


def test_range_tuple(book: Book):
    from xlviews.common import get_range

    cells = get_range((4, 6), book=book)

    assert cells.row == 4
    assert cells.column == 6
    assert cells.shape == (1, 1)


def test_range_str_tuple(book: Book):
    from xlviews.common import get_range

    cells = get_range("get_range_str_tuple", (6, 8), book=book)

    assert cells.row == 6
    assert cells.column == 8
    assert cells.shape == (1, 1)
    assert cells.sheet.name == "get_range_str_tuple"


def test_range_str_str(book: Book):
    from xlviews.common import get_range

    cells = get_range("get_range_str_str", "E12:H20", book=book)

    assert cells.row == 12
    assert cells.column == 5
    assert cells.shape == (9, 4)
    assert cells.sheet.name == "get_range_str_str"


def test_range_int_int(book: Book):
    from xlviews.common import get_range

    cells = get_range(30, 40, book=book)

    assert cells.row == 30
    assert cells.column == 40
    assert cells.shape == (1, 1)


def test_range_str_int_int(book: Book):
    from xlviews.common import get_range

    cells = get_range("get_range_str_int_int", 30, 40, book=book)

    assert cells.row == 30
    assert cells.column == 40
    assert cells.shape == (1, 1)
    assert cells.sheet.name == "get_range_str_int_int"


def test_range_error():
    from xlviews.common import get_range

    with pytest.raises(ValueError, match="Invalid number of arguments: 4"):
        get_range("", 1, 2, 3)


def test_create_book(app: App, tmp_path: Path):
    file = tmp_path / "create_book.xlsx"
    book = open_or_create(file, app, visible=False)

    assert file.exists()
    assert isinstance(book, Book)
    assert book.name == "create_book.xlsx"
    assert len(book.sheets) == 1
    assert book.sheets[0].name == "Sheet1"

    book.close()


def test_create_sheet(app: App, tmp_path: Path):
    file = tmp_path / "create_sheet.xlsx"
    sheet = open_or_create(file, app, sheet_name="New", visible=False)

    assert isinstance(sheet, Sheet)
    assert sheet.name == "New"
    assert sheet.book.name == "create_sheet.xlsx"
    assert len(sheet.book.sheets) == 1

    sheet.book.close()


def test_open_book(app: App, tmp_path: Path):
    file = tmp_path / "open_book.xlsx"
    sheet = open_or_create(file, app, sheet_name="New", visible=False)
    sheet.book.close()

    book = open_or_create(file, app, visible=False)

    assert book.name == "open_book.xlsx"
    assert len(book.sheets) == 1
    assert book.sheets[0].name == "New"

    book.close()


def test_open_sheet(app: App, tmp_path: Path):
    file = tmp_path / "open_sheet.xlsx"
    sheet = open_or_create(file, app, sheet_name="New", visible=False)
    sheet.book.close()
    sheet = open_or_create(file, app, sheet_name="New", visible=False)

    assert sheet.name == "New"
    assert sheet.book.name == "open_sheet.xlsx"
    assert len(sheet.book.sheets) == 1

    sheet.book.close()


def test_create_sheet_after_existing(app: App, tmp_path: Path):
    file = tmp_path / "create_sheet_after_existing.xlsx"
    sheet = open_or_create(file, app, sheet_name="Old", visible=False)
    sheet.book.close()
    sheet = open_or_create(file, app, sheet_name="New", visible=False)

    assert sheet.name == "New"
    assert sheet.book.name == "create_sheet_after_existing.xlsx"
    assert len(sheet.book.sheets) == 2
    assert sheet.book.sheets[0].name == "Old"
    assert sheet.book.sheets[1].name == "New"

    sheet.book.close()
