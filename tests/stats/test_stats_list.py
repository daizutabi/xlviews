import numpy as np
import pytest

from xlviews.frame import SheetFrame
from xlviews.stats import StatsFrame


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    stats = ["count", "max", "median", "soa"]
    return StatsFrame(sf_parent, by=":y", stats=stats, table=True)


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
    sf.auto_filter(func)
    df = sf.visible_data
    assert len(df) == n


@pytest.mark.parametrize(
    ("func", "column", "value"),
    [
        ("count", "a", [5, 4, 6, 3]),
        ("count", "b", [6, 4, 6, 4]),
        ("count", "c", [5, 2, 6, 4]),
        ("max", "a", [5, 9, 15, 18]),
        ("max", "b", [5, 9, 15, 27]),
        ("max", "c", [30, 36, 10, 18]),
        ("median", "a", [2, 7.5, 12.5, 17]),
        ("median", "b", [2.5, 7.5, 7.5, 22.5]),
        ("median", "c", [24, 35, 5, 15]),
    ],
)
def test_value_float(sf: StatsFrame, func, column, value):
    sf.auto_filter(func)
    df = sf.visible_data
    np.testing.assert_allclose(df[column], value)


@pytest.mark.parametrize(
    ("column", "value"),
    [
        ("a", [[0, 1, 2, 3, 5], [6, 7, 8, 9]]),
        ("b", [[0, 1, 2, 3, 4, 5], [6, 7, 8, 9]]),
        ("c", [[20, 22, 24, 28, 30], [34, 36]]),
    ],
)
def test_value_soa(sf: StatsFrame, column, value):
    sf.auto_filter("soa")
    df = sf.visible_data
    soa = [np.std(x) / np.median(x) for x in value]
    np.testing.assert_allclose(df[column].iloc[:2], soa)
