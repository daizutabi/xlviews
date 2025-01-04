import pytest
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


def test_border_zero(sheet: Sheet):
    from xlviews.style import set_border_line

    set_border_line(sheet["C3:D5"], "xlInsideVertical", weight=0, color="red")
    assert sheet["C3:D5"].api.Borders(11).Weight == 2
    assert sheet["C3:D5"].api.Borders(12).Weight == 2


def test_fill(sheet: Sheet):
    from xlviews.style import set_fill

    set_fill(sheet["C3:E5"], color="pink")
    assert sheet["C3:E5"].api.Interior.Color == 13353215


def test_font(sheet: Sheet):
    from xlviews.style import set_font

    rng = sheet["C3"]
    rng.value = "abc"
    set_font(rng, "Times", size=24, bold=True, italic=True, color="green")
    assert rng.api.Font.Name == "Times"
    assert rng.api.Font.Size == 24
    assert rng.api.Font.Bold == 1
    assert rng.api.Font.Italic == 1
    assert rng.api.Font.Color == 32768


def test_font_with_name(sheet: Sheet):
    from xlviews.config import rcParams
    from xlviews.style import set_font

    rng = sheet["C3"]
    rng.value = "abc"
    set_font(rng)
    assert rng.api.Font.Name == rcParams["chart.font.name"]


@pytest.mark.parametrize(
    ("align", "value"),
    [("right", -4152), ("left", -4131), ("center", -4108)],
)
def test_alignment_horizontal(sheet: Sheet, align, value):
    from xlviews.style import set_alignment

    rng = sheet["C3"]
    rng.value = "a"
    set_alignment(rng, horizontal_alignment=align)
    assert rng.api.HorizontalAlignment == value


@pytest.mark.parametrize(
    ("align", "value"),
    [("top", -4160), ("bottom", -4107), ("center", -4108)],
)
def test_alignment_vertical(sheet: Sheet, align, value):
    from xlviews.style import set_alignment

    rng = sheet["C3"]
    rng.value = "a"
    set_alignment(rng, vertical_alignment=align)
    assert rng.api.VerticalAlignment == value


@pytest.mark.parametrize(
    ("axis", "even_color", "odd_color"),
    [(0, 100, 200), (1, 300, 400)],
)
def test_banding(sheet: Sheet, axis, even_color, odd_color):
    from xlviews.style import set_banding

    rng = sheet["C3:F6"]
    set_banding(rng, axis, even_color, odd_color)
    assert rng.api.FormatConditions(1).Interior.Color == even_color
    assert rng.api.FormatConditions(2).Interior.Color == odd_color


def test_hide_succession(sheet: Sheet):
    from xlviews.style import hide_succession

    rng = sheet["C3:C8"]
    rng.options(transpose=True).value = [1, 1, 2, 2, 3, 3]
    rng = sheet["D3:D8"]
    rng.options(transpose=True).value = [1, 1, 1, 2, 2, 2]
    rng = sheet["C3:D8"]
    hide_succession(rng, color="red")
    assert rng.api.FormatConditions(1).Font.Color == 255


def test_hide_unique(sheet: Sheet):
    from xlviews.style import hide_unique

    rng = sheet["C3:C8"]
    rng.options(transpose=True).value = [1, 1, 2, 2, 3, 3]
    rng = sheet["D3:D8"]
    rng.options(transpose=True).value = [1, 1, 1, 1, 1, 1]
    rng = sheet["C2:D2"]
    rng.value = ["a", "b"]
    hide_unique(rng, 6, color="red")
    assert rng.api.FormatConditions(1).Font.Color == 255


def test_hide_gridlines(sheet: Sheet):
    from xlviews.style import hide_gridlines

    hide_gridlines(sheet)
    assert sheet.book.app.api.ActiveWindow.DisplayGridlines is False


@pytest.mark.parametrize(
    "name",
    ["index.name", "index", "columns.name", "columns", "values"],
)
@pytest.mark.parametrize("gray", [True, False])
def test_set_style(sheet: Sheet, name, gray):
    from xlviews.config import rcParams
    from xlviews.style import _set_style
    from xlviews.utils import rgb

    rng = sheet["C3:E5"]
    _set_style(rng[0], rng[-1], name, gray=gray)
    param = f"frame.{name}.fill.color"
    color = rgb("#eeeeee") if gray and name != "values" else rgb(rcParams[param])
    assert rng.api.Interior.Color == color


def a():
    import xlwings as xw

    book = xw.Book()
    sheet = book.sheets[0]
    sheet
