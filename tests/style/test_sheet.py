import pytest
from pandas import DataFrame, Series
from xlwings import Sheet
from xlwings.constants import BordersIndex


@pytest.mark.parametrize(("index", "value"), [("Vertical", 11), ("Horizontal", 12)])
def test_border_index(index, value):
    assert getattr(BordersIndex, f"xlInside{index}") == value


@pytest.mark.parametrize(
    ("weight", "value"),
    [((1, 2, 3, 4), (1, 2, -4138, 4)), (2, (2, 2, 2, 2))],
)
def test_border_edge(sheet: Sheet, weight, value):
    from xlviews.style import set_border_edge

    set_border_edge(sheet["C3:E5"], weight, color="red")
    assert sheet["B3:C5"].api.Borders(11).Weight == value[0]
    assert sheet["B3:C5"].api.Borders(11).Color == 255
    assert sheet["E3:F5"].api.Borders(11).Weight == value[1]
    assert sheet["E3:F5"].api.Borders(11).Color == 255
    assert sheet["C2:E3"].api.Borders(12).Weight == value[2]
    assert sheet["C2:E3"].api.Borders(12).Color == 255
    assert sheet["C5:E6"].api.Borders(12).Weight == value[3]
    assert sheet["C5:E6"].api.Borders(12).Color == 255


def test_border_inside(sheet: Sheet):
    from xlviews.style import set_border_inside

    set_border_inside(sheet["C3:E5"], weight=2, color="red")
    assert sheet["C3:E5"].api.Borders(11).Weight == 2
    assert sheet["C3:E5"].api.Borders(11).Color == 255
    assert sheet["C3:E5"].api.Borders(12).Weight == 2
    assert sheet["C3:E5"].api.Borders(12).Color == 255


def test_border(sheet: Sheet):
    from xlviews.style import set_border

    set_border(sheet["C3:E5"], edge_weight=2, inside_weight=1)
    assert sheet["B3:C5"].api.Borders(11).Weight == 2
    assert sheet["C3:E5"].api.Borders(11).Weight == 1


def a():
    import xlwings as xw

    book = xw.Book()
    sheet = book.sheets[0]
