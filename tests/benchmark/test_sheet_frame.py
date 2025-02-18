import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.core.address import index_to_column_name
from xlviews.dataframes.groupby import GroupBy
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


def create_data_frame(rows: int = 10, columns: int = 20) -> DataFrame:
    values = np.arange(rows * columns).reshape((rows, columns))
    cnames = [index_to_column_name(i + 1) for i in range(columns)]
    df = DataFrame(values, columns=cnames)
    return df.set_index(df.columns.to_list()[:10])


def create_sheet_frame(df: DataFrame, sheet: Sheet) -> SheetFrame:
    return SheetFrame(2, 3, data=df, sheet=sheet)


@pytest.mark.parametrize(("rows", "columns"), [(10, 20), (100, 100), (1000, 100)])
def test_create_sheet_frame(benchmark, sheet: Sheet, rows: int, columns: int):
    df = create_data_frame(rows, columns)
    sf = benchmark(create_sheet_frame, df, sheet)
    assert isinstance(sf, SheetFrame)


@pytest.fixture(
    params=[(100, 20), (1000, 20), (10000, 20), (10, 100), (10, 1000)],
    ids=lambda x: "_".join([str(i) for i in x]),
)
def shape(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture
def sf(shape: tuple[int, int], sheet: Sheet):
    rows, columns = shape
    df = create_data_frame(rows, columns)
    return create_sheet_frame(df, sheet)


@pytest.fixture(
    params=[
        ["A"],
        ["A", "B", "C", "D"],
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
    ],
    ids=lambda x: f"C{len(x)}",
)
def columns(request: pytest.FixtureRequest):
    return request.param


@pytest.mark.parametrize("axis", [0, 1])
def test_ranges(benchmark, sf: SheetFrame, shape, axis):
    x = benchmark(lambda: list(sf.iter_ranges(axis)))
    assert len(x) == shape[1 - axis] - 10 * (1 - axis)


def test_agg(benchmark, sf: SheetFrame, columns):
    x = benchmark(lambda: sf.agg(["sum", "count"], columns))
    assert isinstance(x, DataFrame)
    assert x.shape == (2, len(columns))


def test_groupby(benchmark, sf: SheetFrame, columns):
    x = benchmark(lambda: sf.groupby(columns))
    assert isinstance(x, GroupBy)
    assert len(x) == len(sf)


def test_groupby_agg(benchmark, sf: SheetFrame, columns, shape):
    x = benchmark(lambda: sf.groupby(columns).agg(["sum", "count"]))
    assert isinstance(x, DataFrame)
    assert x.shape == (len(sf), 2 * (shape[1] - 10))
