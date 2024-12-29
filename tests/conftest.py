import pytest
import xlwings as xw


@pytest.fixture(scope="session", autouse=True)
def setup_teardown():
    yield
    for app in xw.apps:
        app.quit()
