import numpy as np
import pytest

from xlviews.sheetframe import SheetFrame
from xlviews.statsframe import StatsFrame


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    funcs = ["count", "max", "median", "soa"]
    return StatsFrame(sf_parent, funcs, by=":y", table=True)


def test_len(sf: StatsFrame):
    assert len(sf) == 16


def test_columns(sf: StatsFrame):
    assert sf.columns == ["func", "x", "y", "z", "a", "b", "c"]


def test_index_columns(sf: StatsFrame):
    assert sf.index_columns == ["func", "x", "y", "z"]


def test_value_columns(sf: StatsFrame):
    assert sf.value_columns == ["a", "b", "c"]


@pytest.mark.parametrize(
    ("func", "n"),
    [
        ("median", 4),
        (["soa"], 4),
        (["count", "max", "median"], 12),
        (["count", "max", "median", "soa"], 16),
    ],
)
def test_value_len(sf: StatsFrame, func, n):
    assert sf.table
    sf.table.auto_filter("func", func)
    df = sf.visible_data
    assert len(df) == n


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
def test_value_float(sf: StatsFrame, func, column, value):
    assert sf.table
    sf.table.auto_filter("func", func)
    df = sf.visible_data
    np.testing.assert_allclose(df[column], value)


@pytest.mark.parametrize(
    ("column", "value"),
    [
        ("a", [[0, 1, 2, 3, 16, 17, 18], [5, 6, 7]]),
        ("b", [[0, 1, 2, 3, 18, 21, 24, 27], [4, 5, 6, 7]]),
        ("c", [[20, 22, 24, 12, 14, 16, 18], [28, 30, 34]]),
    ],
)
def test_value_soa(sf: StatsFrame, column, value):
    assert sf.table
    sf.table.auto_filter("func", "soa")
    df = sf.visible_data
    soa = [np.std(x) / np.median(x) for x in value]
    np.testing.assert_allclose(df[column].iloc[:2], soa)
