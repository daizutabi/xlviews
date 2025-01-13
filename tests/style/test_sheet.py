import pytest
from pandas import DataFrame, MultiIndex
from xlwings import Sheet
from xlwings.constants import BordersIndex

from xlviews.config import rcParams
from xlviews.sheetframe import SheetFrame
from xlviews.utils import rgb


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
    from xlviews.style import _set_style

    rng = sheet["C3:E5"]
    _set_style(rng[0], rng[-1], name, gray=gray)
    param = f"frame.{name}.fill.color"
    color = rgb("#eeeeee") if gray and name != "values" else rgb(rcParams[param])
    assert rng.api.Interior.Color == color


@pytest.fixture(scope="module")
def sf_basic(sheet_module: Sheet):
    from xlviews.style import set_frame_style

    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(2, 2, data=df, style=False, sheet=sheet_module)
    set_frame_style(sf, autofit=True)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [("B2", "index.name"), ("B6", "index"), ("D2", "columns"), ("D5", "values")],
)
def test_frame_style_basic(sf_basic: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.{name}.fill.color"])
    assert sf_basic.sheet[cell].api.Interior.Color == c


def test_frame_style_banding_succession(sheet_module: Sheet):
    from xlviews.style import set_frame_style

    df = DataFrame({"x": [1, 1, 2, 2], "a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    df = df.set_index("x")
    sf = SheetFrame(2, 6, data=df, style=False, sheet=sheet_module)
    set_frame_style(sf, autofit=True, banding=True, succession=True)
    assert sf.sheet["F4"].api.FormatConditions(1)
    assert sf.sheet["H5"].api.FormatConditions(1)


@pytest.fixture(scope="module")
def df_mc():
    df_mc = DataFrame([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]])
    mi = [("a", "b", "c"), ("d", "e", "f"), ("g", "h", "i")]
    df_mc.columns = MultiIndex.from_tuples(mi)
    df_mc.columns.names = ["x", "y", "z"]
    return df_mc


@pytest.fixture(scope="module")
def sf_mc(sheet_module: Sheet, df_mc: DataFrame):
    from xlviews.style import set_frame_style

    sf = SheetFrame(9, 2, data=df_mc, style=False, sheet=sheet_module)
    set_frame_style(sf, autofit=True)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [
        ("B9", "index.name"),
        ("B11", "index.name"),
        ("B12", "index"),
        ("B15", "index"),
        ("C9", "columns.name"),
        ("E10", "columns.name"),
        ("C11", "columns"),
        ("E11", "columns"),
        ("C12", "values"),
        ("E15", "values"),
    ],
)
def test_frame_style_multi_columns(sf_mc: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.{name}.fill.color"])
    assert sf_mc.sheet[cell].api.Interior.Color == c


@pytest.fixture(scope="module")
def df_mic(df_mc: DataFrame):
    df_mic = df_mc.copy()
    df_mic.columns = MultiIndex.from_tuples([("a", "b"), ("c", "d"), ("e", "f")])
    df_mic.columns.names = ["x", "y"]
    i = [("i", "j"), ("k", "l"), ("m", "n"), ("o", "p")]
    df_mic.index = MultiIndex.from_tuples(i)
    return df_mic


@pytest.fixture(scope="module")
def sf_mic(sheet_module: Sheet, df_mic: DataFrame):
    from xlviews.style import set_frame_style

    sf = SheetFrame(9, 7, data=df_mic, style=False, sheet=sheet_module)
    set_frame_style(sf, autofit=True)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [
        ("G9", "index.name"),
        ("H10", "index.name"),
        ("G11", "index"),
        ("H14", "index"),
        ("I9", "columns.name"),
        ("K9", "columns.name"),
        ("I10", "columns"),
        ("K10", "columns"),
        ("I11", "values"),
        ("K14", "values"),
    ],
)
def test_frame_style_multi_index_columns(sf_mic: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.{name}.fill.color"])
    assert sf_mic.sheet[cell].api.Interior.Color == c


@pytest.fixture(scope="module")
def sf_wide(sheet_module: Sheet):
    from xlviews.style import set_wide_column_style

    df = DataFrame({"x": ["i", "j"], "y": ["k", "l"], "a": [1, 2], "b": [3, 4]})
    sf = SheetFrame(24, 2, data=df, style=False, sheet=sheet_module)
    sf.add_wide_column("u", range(3), autofit=True)
    sf.add_wide_column("v", range(4), autofit=True)
    set_wide_column_style(sf, gray=False)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [("G23", ".name"), ("M23", ".name"), ("G24", ""), ("M24", "")],
)
def test_frame_style_wide(sf_wide: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.wide-columns{name}.fill.color"])
    assert sf_wide.sheet[cell].api.Interior.Color == c


def test_table_style(sheet_module: Sheet):
    from xlviews.style import set_table_style

    df = DataFrame({"x": [1, 1, 2, 2], "a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    df = df.set_index("x")
    sf = SheetFrame(17, 2, data=df, style=False, sheet=sheet_module)
    table = sf.as_table(style=False)
    set_table_style(table)
    assert table.sheet.book.api.TableStyles("xlviews")
