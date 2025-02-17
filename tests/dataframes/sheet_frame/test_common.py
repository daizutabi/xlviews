import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture
def sf(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(3, 3, df, sheet)


def test_number_format_kwargs(sf: SheetFrame):
    sf.number_format(autofit=False, a="0", b="0.0")
    assert sf.sheet.range("D4:D7").number_format == "0"
    assert sf.sheet.range("E4:E7").number_format == "0.0"


def test_number_format_dict(sf: SheetFrame):
    sf.number_format({r"[ab]": "0.00"}, autofit=True)
    assert sf.get_number_format("a") == "0.00"
    assert sf.get_number_format("b") == "0.00"


@pytest.mark.parametrize(
    ("alignment", "value"),
    [("left", -4131), ("center", -4108), ("right", -4152)],
)
def test_alignment(sf: SheetFrame, alignment: str, value: int):
    sf.alignment(alignment)
    assert sf.cell.api.HorizontalAlignment == value


def test_adjacent_column_width(sf: SheetFrame):
    sf.set_adjacent_column_width(10)
    assert sf.sheet["F1"].column_width == 10


# def test_child_frame(sf: SheetFrame):
#     cell = sf.get_child_cell()
#     assert cell.get_address() == "$G$3"

#     cell = sf.get_adjacent_cell(offset=3)
#     assert cell.get_address() == "$J$3"

#     df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
#     sf_child = SheetFrame(parent=sf, data=df)
#     assert sf_child.cell.get_address() == "$G$3"

#     assert sf_child.parent is sf
#     assert sf.children[0] is sf_child

#     cell = sf.get_adjacent_cell()
#     assert cell.get_address() == "$L$3"


# def test_head_frame(sf: SheetFrame):
#     df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
#     sf_tail = SheetFrame(head=sf, data=df)
#     assert sf_tail.cell.get_address() == "$C$9"
#     assert sf.tail is sf_tail
#     assert sf_tail.head is sf


def test_active_sheet(sheet: Sheet):
    sheet.name = "active"
    sheet.activate()
    df = DataFrame({"x": [1, 2], "y": [5, 6], "z": [7, 8]})
    sf = SheetFrame(100, 10, data=df)
    assert sf.sheet.name == "active"
    assert sf.cell.get_address() == "$J$100"
