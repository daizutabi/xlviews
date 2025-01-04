import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture
def sf(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(sheet, 3, 3, data=df, style=False)


@pytest.mark.parametrize("number_format", ["0", "0.00", "0.00%"])
def test_number_format(sf: SheetFrame, number_format: str):
    sf.set_number_format(number_format, autofit=True)
    assert sf.range("a", -1).number_format == number_format
    assert sf.range("b", -1).number_format == number_format


def test_number_format_kwargs(sf: SheetFrame):
    sf.set_number_format(autofit=False, a="0", b="0.0")
    assert sf.range("a", -1).number_format == "0"
    assert sf.range("b", -1).number_format == "0.0"


def test_number_format_dict(sf: SheetFrame):
    sf.set_number_format({r"[ab]": "0.00"}, autofit=True)
    assert sf.get_number_format("a") == "0.00"
    assert sf.get_number_format("b") == "0.00"


def test_style_gray(sf: SheetFrame):
    sf.add_wide_column("u", range(3), autofit=True)
    sf.set_style(gray=True)
    assert sf.sheet["F2"].api.Font.Bold
    assert sf.sheet["D3"].api.Font.Bold
    assert sf.sheet["C3"].api.Interior.Color == 15658734
    assert sf.sheet["H2"].api.Interior.Color == 15658734
    assert sf.sheet["H7"].api.Interior.Color != 15658734
