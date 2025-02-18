import pytest
from xlwings import Sheet
from xlwings.constants import BordersIndex

from xlviews.config import rcParams
from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize(("index", "value"), [("Vertical", 11), ("Horizontal", 12)])
def test_border_index(index, value):
    assert getattr(BordersIndex, f"xlInside{index}") == value


@pytest.fixture(params=[lambda x: x, Range.from_range], ids=["RangeImpl", "Range"])
def cls(request: pytest.FixtureRequest):
    return request.param


@pytest.mark.parametrize(
    ("weight", "value"),
    [((1, 2, 3, 4), (1, 2, -4138, 4)), (2, (2, 2, 2, 2))],
)
def test_border_edge(weight, value, cls, sheet: Sheet):
    from xlviews.style import set_border_edge

    set_border_edge(cls(sheet["C3:E5"]), weight, color="red")
    assert sheet["B3:C5"].api.Borders(11).Weight == value[0]
    assert sheet["B3:C5"].api.Borders(11).Color == 255
    assert sheet["E3:F5"].api.Borders(11).Weight == value[1]
    assert sheet["E3:F5"].api.Borders(11).Color == 255
    assert sheet["C2:E3"].api.Borders(12).Weight == value[2]
    assert sheet["C2:E3"].api.Borders(12).Color == 255
    assert sheet["C5:E6"].api.Borders(12).Weight == value[3]
    assert sheet["C5:E6"].api.Borders(12).Color == 255


def test_border_inside(cls, sheet: Sheet):
    from xlviews.style import set_border_inside

    set_border_inside(cls(sheet["C3:E5"]), weight=2, color="red")
    assert sheet["C3:E5"].api.Borders(11).Weight == 2
    assert sheet["C3:E5"].api.Borders(11).Color == 255
    assert sheet["C3:E5"].api.Borders(12).Weight == 2
    assert sheet["C3:E5"].api.Borders(12).Color == 255


def test_border(cls, sheet: Sheet):
    from xlviews.style import set_border

    set_border(cls(sheet["C3:E5"]), edge_weight=2, inside_weight=1)
    assert sheet["B3:C5"].api.Borders(11).Weight == 2
    assert sheet["C3:E5"].api.Borders(11).Weight == 1


def test_border_zero(cls, sheet: Sheet):
    from xlviews.style import set_border_line

    set_border_line(cls(sheet["C3:D5"]), "xlInsideVertical", weight=0, color="red")
    assert sheet["C3:D5"].api.Borders(11).Weight == 2
    assert sheet["C3:D5"].api.Borders(12).Weight == 2


def test_fill(cls, sheet: Sheet):
    from xlviews.style import set_fill

    set_fill(cls(sheet["C3:E5"]), color="pink")
    assert sheet["C3:E5"].api.Interior.Color == 13353215


def test_font(cls, sheet: Sheet):
    from xlviews.style import set_font

    rng = cls(sheet["C3"])
    rng.value = "abc"
    set_font(rng, "Times", size=24, bold=True, italic=True, color="green")
    assert rng.api.Font.Name == "Times"
    assert rng.api.Font.Size == 24
    assert rng.api.Font.Bold == 1
    assert rng.api.Font.Italic == 1
    assert rng.api.Font.Color == 32768


def test_font_collection(sheet: Sheet):
    from xlviews.core.range_collection import RangeCollection
    from xlviews.style import set_font

    rc = RangeCollection([(2, 3), (6, 7)], 2, sheet)
    set_font(rc, "Times", size=24, bold=True, italic=True, color="green")

    for row in [2, 3, 6, 7]:
        rng = sheet.range(row, 2)
        assert rng.api.Font.Name == "Times"
        assert rng.api.Font.Size == 24
        assert rng.api.Font.Bold == 1
        assert rng.api.Font.Italic == 1
        assert rng.api.Font.Color == 32768

    for row in [4, 5]:
        rng = sheet.range(row, 2)
        assert rng.api.Font.Size != 24


def test_font_without_name(cls, sheet: Sheet):
    from xlviews.style import set_font

    rng = cls(sheet["C3"])
    rng.value = "abc"
    set_font(rng)
    assert rng.api.Font.Name == rcParams["frame.font.name"]


@pytest.mark.parametrize(
    ("align", "value"),
    [("right", -4152), ("left", -4131), ("center", -4108)],
)
def test_alignment_horizontal(align, value, cls, sheet: Sheet):
    from xlviews.style import set_alignment

    rng = cls(sheet["C3"])
    rng.value = "a"
    set_alignment(rng, horizontal_alignment=align)
    assert rng.api.HorizontalAlignment == value


@pytest.mark.parametrize(
    ("align", "value"),
    [("top", -4160), ("bottom", -4107), ("center", -4108)],
)
def test_alignment_vertical(align, value, cls, sheet: Sheet):
    from xlviews.style import set_alignment

    rng = cls(sheet["C3"])
    rng.value = "a"
    set_alignment(rng, vertical_alignment=align)
    assert rng.api.VerticalAlignment == value


def test_number_format(cls, sheet: Sheet):
    from xlviews.style import set_number_format

    rng = cls(sheet["C3"])
    set_number_format(rng, "0.0%")
    assert rng.api.NumberFormat == "0.0%"


def test_number_format_collection(sheet: Sheet):
    from xlviews.core.range_collection import RangeCollection
    from xlviews.style import set_number_format

    rc = RangeCollection([(2, 3), (6, 7)], 3, sheet)
    set_number_format(rc, "0.0%")

    for row in [2, 3, 6, 7]:
        rng = sheet.range(row, 3)
        assert rng.api.NumberFormat == "0.0%"

    for row in [4, 5]:
        rng = sheet.range(row, 3)
        assert rng.api.NumberFormat != "0.0%"


@pytest.mark.parametrize(
    ("axis", "even_color", "odd_color"),
    [(0, 100, 200), (1, 300, 400)],
)
def test_banding(axis, even_color, odd_color, cls, sheet: Sheet):
    from xlviews.style import set_banding

    rng = cls(sheet["C3:F6"])
    set_banding(rng, axis, even_color, odd_color)
    assert rng.api.FormatConditions(1).Interior.Color == even_color
    assert rng.api.FormatConditions(2).Interior.Color == odd_color


def test_hide_succession(cls, sheet: Sheet):
    from xlviews.style import hide_succession

    rng = sheet["C3:C8"]
    rng.options(transpose=True).value = [1, 1, 2, 2, 3, 3]
    rng = sheet["D3:D8"]
    rng.options(transpose=True).value = [1, 1, 1, 2, 2, 2]
    rng = cls(sheet["C3:D8"])
    hide_succession(rng, color="red")
    assert rng.api.FormatConditions(1).Font.Color == 255


def test_hide_unique(cls, sheet: Sheet):
    from xlviews.style import hide_unique

    rng = sheet["C3:C8"]
    rng.options(transpose=True).value = [1, 1, 2, 2, 3, 3]
    rng = sheet["D3:D8"]
    rng.options(transpose=True).value = [1, 1, 1, 1, 1, 1]
    rng = cls(sheet["C2:D2"])
    rng.value = ["a", "b"]
    hide_unique(rng, 6, color="red")
    assert rng.api.FormatConditions(1).Font.Color == 255


def test_hide_gridlines(sheet: Sheet):
    from xlviews.style import hide_gridlines

    hide_gridlines(sheet)
    assert sheet.book.app.api.ActiveWindow.DisplayGridlines is False
