import numpy as np
import pytest
from pandas import DataFrame

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.stats_frame import StatsFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    funcs = ["count", "max", "median", "soa"]
    sf = StatsFrame(sf_parent, funcs, by=":y")
    sf.as_table()
    return sf


@pytest.fixture(scope="module")
def df(sf: StatsFrame):
    rng = sf.expand().impl
    return rng.options(DataFrame, index=sf.index.nlevels).value


def test_len(sf: StatsFrame):
    assert len(sf) == 16


def test_index_names(sf: StatsFrame):
    assert sf.index.names == ["func", "x", "y", "z"]


@pytest.mark.parametrize(
    ("func", "column", "value"),
    [
        ("count", "a", [7, 3, 4, 4]),
        ("count", "b", [8, 4, 4, 4]),
        ("count", "c", [7, 3, 3, 4]),
        ("max", "a", [18, 7, 11, 15]),
        ("max", "b", [27, 7, 9, 15]),
        ("max", "c", [24, 34, 36, 10]),
        ("median", "a", [3, 6, 9.5, 13.5]),
        ("median", "b", [10.5, 5.5, 5.5, 10.5]),
        ("median", "c", [18, 30, 2, 7]),
    ],
)
def test_value_float(df: DataFrame, func, column, value):
    np.testing.assert_allclose(df.loc[func][column], value)


@pytest.mark.parametrize(
    ("column", "value"),
    [
        ("a", [[0, 1, 2, 3, 16, 17, 18], [5, 6, 7]]),
        ("b", [[0, 1, 2, 3, 18, 21, 24, 27], [4, 5, 6, 7]]),
        ("c", [[20, 22, 24, 12, 14, 16, 18], [28, 30, 34]]),
    ],
)
def test_value_soa(df: DataFrame, column, value):
    soa = [np.std(x) / np.median(x) for x in value]
    np.testing.assert_allclose(df.loc["soa"][column][:2], soa)
