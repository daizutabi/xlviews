import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews import SheetFrame
from xlviews.testing import is_excel_installed
from xlviews.testing.sheet_frame.pivot import Base

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return Base(sheet_module)


@pytest.fixture(scope="module")
def df_parent(fc_parent: Base):
    return fc_parent.df


@pytest.fixture(scope="module")
def sf_parent(fc_parent: Base):
    return fc_parent.sf


@pytest.fixture(scope="module", params=["u", "v", None])
def values(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module", params=["y", ["y"]])
def index(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module", params=["x", ["x"]])
def columns(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame, values, index, columns):
    df = sf_parent.pivot_table(values, index, columns, formula=True)
    return SheetFrame(100 if values is None else 80, 2, df)


@pytest.fixture(scope="module")
def df(df_parent: DataFrame, values, index, columns):
    return df_parent.pivot_table(values, index, columns, aggfunc=lambda x: x)


def test_index(sf: SheetFrame, df: DataFrame):
    assert sf.value.index.equals(df.index)


def test_columns(sf: SheetFrame, df: DataFrame):
    assert sf.value.columns.equals(df.columns)


def test_values(sf: SheetFrame, df: DataFrame):
    assert sf.value.equals(df)
