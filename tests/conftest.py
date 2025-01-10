import pytest
import xlwings
from xlwings import App, Book


@pytest.fixture(scope="session", autouse=True)
def teardown():
    from xlviews.common import quit_apps

    yield
    quit_apps()


@pytest.fixture(scope="session")
def app():
    return xlwings.apps.add()

    # app.quit() # Avoids quitting the app to execute `quit_apps()` later


@pytest.fixture(scope="session")
def book(app: App):
    return app.books.add()


@pytest.fixture(scope="module")
def sheet_module(book: Book):
    from xlviews.style import hide_gridlines

    sheet = book.sheets.add()
    hide_gridlines(sheet)

    yield sheet

    sheet.delete()


@pytest.fixture
def sheet(book: Book):
    sheet = book.sheets.add()

    yield sheet

    sheet.delete()
