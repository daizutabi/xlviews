import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture
def sf_nf(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(sheet, 2, 3, data=df, style=False)


@pytest.mark.parametrize("number_format", ["0", "0.00", "0.00%"])
def test_number_format(sf_nf: SheetFrame, number_format: str):
    sf_nf.set_number_format(number_format, autofit=True)
    assert sf_nf.range("a", -1).number_format == number_format
    assert sf_nf.range("b", -1).number_format == number_format


def test_number_format_kwargs(sf_nf: SheetFrame):
    sf_nf.set_number_format(autofit=False, a="0", b="0.0")
    assert sf_nf.range("a", -1).number_format == "0"
    assert sf_nf.range("b", -1).number_format == "0.0"


def test_number_format_dict(sf_nf: SheetFrame):
    sf_nf.set_number_format({r"[ab]": "0.00"}, autofit=True)
    assert sf_nf.get_number_format("a") == "0.00"
    assert sf_nf.get_number_format("b") == "0.00"


def a():
    import xlwings as xw

    book = xw.Book()
    sheet = book.sheets[0]
    sheet
