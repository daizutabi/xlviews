import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer, is_app_available
from xlviews.testing.sheet_frame.base import MultiIndex

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return MultiIndex(sheet_module)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: FrameContainer):
    return fc.sf


def test_init(sf: SheetFrame, sheet_module: Sheet):
    assert sf.row == 2
    assert sf.column == 6
    assert sf.sheet.name == sheet_module.name
    assert sf.index.nlevels == 2
    assert sf.columns.nlevels == 1


def test_len(sf: SheetFrame):
    assert len(sf) == 8


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["x", "y"]


@pytest.mark.parametrize(("x", "b"), [("x", False), ("b", True)])
def test_contains(sf: SheetFrame, x, b):
    assert (x in sf) is b


def test_iter(sf: SheetFrame):
    assert list(sf) == ["a", "b"]


def test_value(sf: SheetFrame, df: DataFrame):
    df_sf = sf.value
    assert df_sf.equals(df.astype(float))
    assert df_sf.index.equals(df.index)
    assert df_sf.columns.equals(df.columns)


@pytest.mark.parametrize(
    ("column", "offset", "address"),
    [
        ("x", -1, "$F$2"),
        ("y", 0, "$G$3"),
        ("a", -1, "$H$2"),
        ("b", 0, "$I$3"),
        ("y", None, "$G$3:$G$10"),
    ],
)
def test_get_range(sf: SheetFrame, column, offset, address):
    assert sf.get_range(column, offset).get_address() == address


@pytest.mark.parametrize(
    ("by", "v1", "v2"),
    [
        ("x", [(3, 6)], [(7, 10)]),
        ("y", [(3, 4), (7, 8)], [(5, 6), (9, 10)]),
    ],
)
def test_groupby(sf: SheetFrame, by, v1, v2):
    g = groupby(sf, by)
    assert len(g) == 2
    assert g[(1,)] == v1
    assert g[(2,)] == v2


def test_groupby_list(sf: SheetFrame):
    g = groupby(sf, ["x", "y"])
    assert len(g) == 4
    assert g[(1, 1)] == [(3, 4)]
    assert g[(1, 2)] == [(5, 6)]
    assert g[(2, 1)] == [(7, 8)]
    assert g[(2, 2)] == [(9, 10)]


def test_get_address(sf: SheetFrame):
    df = sf.get_address(row_absolute=False, column_absolute=False, formula=True)
    assert df.columns.to_list() == ["a", "b"]
    assert df.index.names == ["x", "y"]
    assert df.to_numpy()[0, 0] == "=H3"
