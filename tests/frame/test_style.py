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


@pytest.fixture(scope="module")
def sf_wide(sheet_module: Sheet):
    df = DataFrame({"x": ["i", "j"], "y": ["k", "l"], "a": [1, 2], "b": [3, 4]})
    sf = SheetFrame(sheet_module, 4, 2, data=df, style=False)
    sf.add_wide_column("u", range(3), autofit=True)
    sf.add_wide_column("v", range(4), autofit=True)
    sf.set_style(alignment="left", gray=False)
    return sf


@pytest.mark.parametrize(
    ("cell", "color"),
    [("G3", "#e0ffe0"), ("M3", "#e0ffe0"), ("G4", "#f0fff0"), ("M4", "#f0fff0")],
)
def test_frame_style_basic(sf_wide: SheetFrame, cell: str, color: str):
    from xlviews.style import rgb

    c = rgb(color)
    assert sf_wide.sheet[cell].api.Interior.Color == c
