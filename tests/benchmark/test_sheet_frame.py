import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.core.address import index_to_column_name
from xlviews.dataframes.groupby import GroupBy
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def create_data_frame(rows: int = 10, columns: int = 10) -> DataFrame:
    values = np.arange(rows * columns).reshape((rows, columns))
    cnames = [index_to_column_name(i + 1) for i in range(columns)]
    df = DataFrame(values, columns=cnames)
    df.index.name = "name"
    return df


def create_sheet_frame(df: DataFrame, sheet: Sheet) -> SheetFrame:
    return SheetFrame(2, 3, data=df, sheet=sheet)


@pytest.mark.parametrize(("rows", "columns"), [(10, 10), (100, 100), (1000, 100)])
def test_create_sheet_frame(benchmark, sheet: Sheet, rows: int, columns: int):
    df = create_data_frame(rows, columns)
    sf = benchmark(create_sheet_frame, df, sheet)
    assert isinstance(sf, SheetFrame)


@pytest.fixture(
    params=[(100, 10), (1000, 10), (10000, 10), (10, 100), (10, 1000)],
    ids=lambda x: "_".join([str(i) for i in x]),
)
def shape(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture
def sf(shape: tuple[int, int], sheet: Sheet):
    rows, columns = shape
    df = create_data_frame(rows, columns)
    return create_sheet_frame(df, sheet)


def test_len(benchmark, sf: SheetFrame, shape):
    assert benchmark(len, sf) == shape[0]


def test_columns(benchmark, sf: SheetFrame, shape):
    assert benchmark(lambda: len(sf.columns.to_list())) == shape[1] + 1


def test_index_str(benchmark, sf: SheetFrame):
    assert benchmark(lambda: sf.index_past("E")) == 8


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


def test_index_list(benchmark, sf: SheetFrame, columns):
    x = benchmark(lambda: sf.index_past(columns))
    assert x == list(range(4, len(columns) + 4))


def test_range(benchmark, sf: SheetFrame, shape):
    x = benchmark(lambda: sf.range("A"))
    assert len(x) == shape[0]


def test_ranges(benchmark, sf: SheetFrame, shape):
    x = benchmark(lambda: list(sf.ranges()))
    assert len(x) == shape[0]


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
    assert x.shape == (len(sf), 2 * shape[1])
