import numpy as np
import pytest
from pandas import DataFrame, Series
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_excel_installed
from xlviews.utils import int_to_column_name

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def create_data_frame(rows: int = 10, columns: int = 10) -> DataFrame:
    values = np.arange(rows * columns).reshape((rows, columns))
    cnames = [int_to_column_name(i + 1) for i in range(columns)]
    df = DataFrame(values, columns=cnames)
    df.index.name = "name"
    return df


def create_sheet_frame(df: DataFrame, sheet: Sheet) -> SheetFrame:
    return SheetFrame(2, 3, data=df, sheet=sheet, style=False)


@pytest.mark.parametrize(("rows", "columns"), [(10, 10), (100, 100), (1000, 100)])
def test_create_sheet_frame(benchmark, sheet: Sheet, rows: int, columns: int):
    df = create_data_frame(rows, columns)
    sf = benchmark(create_sheet_frame, df, sheet)
    assert isinstance(sf, SheetFrame)


@pytest.fixture(
    scope="module",
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


def test_len(benchmark, sf, shape):
    assert benchmark(len, sf) == shape[0]


def test_columns(benchmark, sf, shape):
    assert benchmark(lambda: len(sf.columns)) == shape[1] + 1


def test_index(benchmark, sf):
    assert benchmark(lambda: sf.index("E")) == 8


def test_range(benchmark, sf, shape):
    x = benchmark(lambda: sf.range("A"))
    assert len(x) == shape[0]


def test_get_address_str(benchmark, sf, shape):
    x = benchmark(lambda: sf.get_address("A"))
    assert isinstance(x, Series)
    assert len(x) == shape[0]


@pytest.mark.parametrize(
    "columns",
    [["A"], ["B", "C", "D"], ["E", "F", "G", "H", "I", "J"]],
    ids=lambda x: len(x),
)
def test_get_address_list(benchmark, sf, columns, shape):
    x = benchmark(lambda: sf.get_address(columns))
    assert isinstance(x, DataFrame)
    assert x.shape == (shape[0], len(columns))


if __name__ == "__main__":
    from xlviews.testing import create_sheet

    sheet = create_sheet()
    df_ = create_data_frame(1000, 100)
    sf_ = create_sheet_frame(df_, sheet)

    # print(sf_.get_address())
