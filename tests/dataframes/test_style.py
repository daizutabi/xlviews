import pytest
from pandas import DataFrame, MultiIndex
from xlwings import Sheet

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize(
    "name",
    ["index.name", "index", "columns.name", "columns", "values"],
)
def test_set_style(sheet: Sheet, name):
    from xlviews.dataframes.style import _set_style

    rng = sheet["C3:E5"]
    _set_style(rng[0], rng[-1], name)
    param = f"frame.{name}.fill.color"
    color = rgb(rcParams[param])
    assert rng.api.Interior.Color == color


@pytest.fixture(scope="module")
def sf_basic(sheet_module: Sheet):
    from xlviews.dataframes.style import set_frame_style

    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    sf = SheetFrame(2, 2, data=df, sheet=sheet_module)
    set_frame_style(sf)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [("B2", "index.name"), ("B6", "index"), ("D2", "columns"), ("D5", "values")],
)
def test_frame_style_basic(sf_basic: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.{name}.fill.color"])
    assert sf_basic.sheet[cell].api.Interior.Color == c


def test_frame_style_banding_succession(sheet_module: Sheet):
    from xlviews.dataframes.style import set_frame_style

    df = DataFrame({"x": [1, 1, 2, 2], "a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    df = df.set_index("x")
    sf = SheetFrame(2, 6, data=df, sheet=sheet_module)
    set_frame_style(sf, banding=True, succession=True)
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
    from xlviews.dataframes.style import set_frame_style

    sf = SheetFrame(9, 2, data=df_mc, sheet=sheet_module)
    set_frame_style(sf)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [
        ("B9", "columns.name"),
        ("B11", "columns.name"),
        ("B12", "index"),
        ("B15", "index"),
        ("C9", "columns"),
        ("E10", "columns"),
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
    from xlviews.dataframes.style import set_frame_style

    sf = SheetFrame(9, 7, data=df_mic, sheet=sheet_module)
    set_frame_style(sf)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [
        ("G9", "index.name"),
        ("H10", "index.name"),
        ("G11", "index"),
        ("H14", "index"),
        ("I9", "columns"),
        ("K9", "columns"),
        ("I10", "columns"),
        ("K10", "columns"),
        ("I11", "values"),
        ("K14", "values"),
    ],
)
def test_frame_style_multi_index_names(sf_mic: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.{name}.fill.color"])
    assert sf_mic.sheet[cell].api.Interior.Color == c


@pytest.fixture(scope="module")
def sf_wide(sheet_module: Sheet):
    from xlviews.dataframes.style import set_wide_column_style

    df = DataFrame({"x": ["i", "j"], "y": ["k", "l"], "a": [1, 2], "b": [3, 4]})
    sf = SheetFrame(24, 2, data=df, sheet=sheet_module)
    sf.add_wide_column("u", range(3), autofit=True)
    sf.add_wide_column("v", range(4), autofit=True)
    set_wide_column_style(sf)
    return sf


@pytest.mark.parametrize(
    ("cell", "name"),
    [("G23", ".name"), ("M23", ".name"), ("G24", ""), ("M24", "")],
)
def test_frame_style_wide(sf_wide: SheetFrame, cell: str, name: str):
    c = rgb(rcParams[f"frame.wide-columns{name}.fill.color"])
    assert sf_wide.sheet[cell].api.Interior.Color == c


def test_table_style(sheet_module: Sheet):
    from xlviews.dataframes.style import set_table_style

    df = DataFrame({"x": [1, 1, 2, 2], "a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    df = df.set_index("x")
    sf = SheetFrame(17, 2, data=df, sheet=sheet_module)
    table = sf.as_table(style=False)
    assert table
    set_table_style(table)
    assert table.sheet.book.api.TableStyles("xlviews")
