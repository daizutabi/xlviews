import pytest
import xlwings
from xlwings import App, Book


@pytest.fixture(scope="session", autouse=True)
def teardown():
    yield
    for app in xlwings.apps:
        app.quit()


@pytest.fixture(scope="session")
def app():
    app = xlwings.apps.add()

    yield app

    app.quit()


@pytest.fixture(scope="session")
def book(app: App):
    return app.books.add()


@pytest.fixture
def sheet(book: Book):
    sheet = book.sheets.add()

    yield sheet

    sheet.delete()
