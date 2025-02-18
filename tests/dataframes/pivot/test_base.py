import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.sheet_frame.pivot import Base

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return Base(sheet_module)


@pytest.fixture(scope="module")
def df_parent(fc_parent: Base):
    return fc_parent.df


@pytest.fixture(scope="module")
def sf_parent(fc_parent: Base):
    return fc_parent.sf


@pytest.fixture(params=["u", "v", None, ["v", "u"]])
def values(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(params=["y", ["y"]])
def index(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(params=["x", ["x"]])
def columns(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture
def sf(sf_parent: SheetFrame, values, index, columns, sheet):
    df = sf_parent.pivot_table(
        values,
        index,
        columns,
        include_sheetname=True,
        formula=True,
    )
    return SheetFrame(2, 2, df, sheet=sheet)


@pytest.fixture
def df(df_parent: DataFrame, values, index, columns):
    return df_parent.pivot_table(values, index, columns, aggfunc=lambda x: x)


def test_index(sf: SheetFrame, df: DataFrame):
    assert sf.value.index.equals(df.index)


def test_columns(sf: SheetFrame, df: DataFrame):
    assert sf.value.columns.equals(df.columns)


def test_values(sf: SheetFrame, df: DataFrame):
    assert sf.value.equals(df)


@pytest.fixture(params=[("y", "x"), ("y", None), (None, "x")])
def index_columns(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(params=["mean", ["sum", "count"]])
def aggfunc(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture
def sf_agg(sf_parent: SheetFrame, values, index_columns, aggfunc, sheet):
    index, columns = index_columns
    df = sf_parent.pivot_table(
        values,
        index,
        columns,
        aggfunc,
        include_sheetname=True,
        formula=True,
    )
    return SheetFrame(2, 2, df, sheet=sheet)


@pytest.fixture
def df_agg(df_parent: DataFrame, values, index_columns, aggfunc):
    index, columns = index_columns
    return df_parent.pivot_table(values, index, columns, aggfunc)


def test_agg_shape(sf_agg: SheetFrame, df_agg: DataFrame):
    assert sf_agg.value.shape == df_agg.shape


def test_agg_index(sf_agg: SheetFrame, df_agg: DataFrame):
    assert sf_agg.value.index.equals(df_agg.index)


def test_agg_columns(sf_agg: SheetFrame, df_agg: DataFrame):
    assert sf_agg.value.columns.equals(df_agg.columns)


def test_agg_values(sf_agg: SheetFrame, df_agg: DataFrame):
    assert sf_agg.value.equals(df_agg.astype(float))


@pytest.mark.parametrize("aggfunc", [None, "mean"])
def test_error(sf_parent: SheetFrame, aggfunc):
    with pytest.raises(ValueError, match="No group keys passed!"):
        sf_parent.pivot_table(None, None, None, aggfunc, formula=True)
