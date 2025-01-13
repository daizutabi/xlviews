import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture
def sf(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(3, 3, data=df, style=False, sheet=sheet)


@pytest.mark.parametrize("number_format", ["0", "0.00", "0.00%"])
def test_number_format(sf: SheetFrame, number_format: str):
    sf.set_number_format(number_format, autofit=True)
    assert sf.range("a").number_format == number_format
    assert sf.range("b").number_format == number_format


def test_number_format_kwargs(sf: SheetFrame):
    sf.set_number_format(autofit=False, a="0", b="0.0")
    assert sf.range("a").number_format == "0"
    assert sf.range("b").number_format == "0.0"


def test_number_format_dict(sf: SheetFrame):
    sf.set_number_format({r"[ab]": "0.00"}, autofit=True)
    assert sf.get_number_format("a") == "0.00"
    assert sf.get_number_format("b") == "0.00"


def test_style_gray(sf: SheetFrame):
    sf.add_wide_column("u", range(3), autofit=True, number_format="0", style=True)
    sf.set_style(gray=True)
    assert sf.sheet["F2"].api.Font.Bold
    assert sf.sheet["D3"].api.Font.Bold
    assert sf.sheet["C3"].api.Interior.Color == 15658734
    assert sf.sheet["H2"].api.Interior.Color == 15658734
    assert sf.sheet["H7"].api.Interior.Color != 15658734


@pytest.mark.parametrize(
    ("alignment", "value"),
    [("left", -4131), ("center", -4108), ("right", -4152)],
)
def test_alignment(sf: SheetFrame, alignment: str, value: int):
    sf.set_alignment(alignment)
    assert sf.cell.api.HorizontalAlignment == value


def test_adjacent_column_width(sf: SheetFrame):
    sf.set_adjacent_column_width(10)
    assert sf.sheet["F1"].column_width == 10


def test_child_frame(sf: SheetFrame):
    cell = sf.get_child_cell()
    assert cell.get_address() == "$G$3"

    cell = sf.get_adjacent_cell(offset=3)
    assert cell.get_address() == "$J$3"

    df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
    sf_child = SheetFrame(parent=sf, data=df, style=False)
    assert sf_child.cell.get_address() == "$G$3"

    assert sf_child.parent is sf
    assert sf.children[0] is sf_child

    cell = sf.get_adjacent_cell()
    assert cell.get_address() == "$L$3"


def test_head_frame(sf: SheetFrame):
    df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
    sf_tail = SheetFrame(head=sf, data=df, number_format="0.00")
    assert sf_tail.cell.get_address() == "$C$9"
    assert sf.tail is sf_tail
    assert sf_tail.head is sf


def test_active_sheet(sheet: Sheet):
    sheet.name = "active"
    sheet.activate()
    df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
    sf = SheetFrame(100, 10, data=df, style=False)
    assert sf.sheet.name == "active"
    assert sf.cell.get_address() == "$J$100"
