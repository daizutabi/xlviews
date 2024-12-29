from pathlib import Path

from xlwings import Book, Sheet

from xlviews.main import open_or_create


def test_create_book(tmp_path: Path):
    path = tmp_path / "create_book.xlsx"
    book = open_or_create(path)

    assert path.exists()
    assert isinstance(book, Book)
    assert book.name == "create_book.xlsx"
    assert len(book.sheets) == 1
    assert book.sheets[0].name == "Sheet1"


def test_create_sheet(tmp_path: Path):
    path = tmp_path / "create_sheet.xlsx"
    sheet = open_or_create(path, sheet_name="New")

    assert isinstance(sheet, Sheet)
    assert sheet.name == "New"
    assert sheet.book.name == "create_sheet.xlsx"
    assert len(sheet.book.sheets) == 1


def test_open_book(tmp_path: Path):
    path = tmp_path / "open_book.xlsx"
    sheet = open_or_create(path, sheet_name="New")
    sheet.book.close()
    book = open_or_create(path)

    assert book.name == "open_book.xlsx"
    assert len(book.sheets) == 1
    assert book.sheets[0].name == "New"


def test_open_sheet(tmp_path: Path):
    path = tmp_path / "open_sheet.xlsx"
    sheet = open_or_create(path, sheet_name="New")
    sheet.book.close()
    sheet = open_or_create(path, sheet_name="New")

    assert sheet.name == "New"
    assert sheet.book.name == "open_sheet.xlsx"
    assert len(sheet.book.sheets) == 1


def test_create_sheet_after_existing(tmp_path: Path):
    path = tmp_path / "create_sheet_after_existing.xlsx"
    sheet = open_or_create(path, sheet_name="Old")
    sheet.book.close()
    sheet = open_or_create(path, sheet_name="New")

    assert sheet.name == "New"
    assert sheet.book.name == "create_sheet_after_existing.xlsx"
    assert len(sheet.book.sheets) == 2
    assert sheet.book.sheets[0].name == "Old"
    assert sheet.book.sheets[1].name == "New"
