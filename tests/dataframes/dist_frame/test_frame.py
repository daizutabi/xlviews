import numpy as np
import pytest
from pandas import Index, MultiIndex
from scipy.stats import norm

from xlviews.dataframes.dist_frame import DistFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


def test_select_index():
    from xlviews.dataframes.dist_frame import select_index

    index = Index(["a", "b"], name="x")
    x = select_index(index, ["x"])
    assert x.equals(index)


def test_select_index_none():
    from xlviews.dataframes.dist_frame import select_index

    index = Index(["a", "b"], name="x")
    x = select_index(index, ["y"])
    assert x.equals(Index([0, 1]))


def test_select_index_multi_one():
    from xlviews.dataframes.dist_frame import select_index

    index = MultiIndex.from_tuples([(1, 2), (3, 4)], names=["x", "y"])
    x = select_index(index, ["x"])
    assert isinstance(x, Index)
    assert not isinstance(x, MultiIndex)
    assert x.name == "x"
    assert x.names == ["x"]


def test_select_index_multi_two():
    from xlviews.dataframes.dist_frame import select_index

    index = MultiIndex.from_tuples([(1, 2), (3, 4)], names=["x", "y"])
    x = select_index(index, ["x", "y"])
    assert isinstance(x, MultiIndex)
    assert x.names == ["x", "y"]


def test_select_index_multi_none():
    from xlviews.dataframes.dist_frame import select_index

    index = MultiIndex.from_tuples([(1, 2), (3, 4)], names=["x", "y"])
    x = select_index(index, [])
    assert x.equals(Index([0, 1]))


def test_init_data(sf_parent: SheetFrame):
    from xlviews.dataframes.dist_frame import get_init_data

    df = get_init_data(sf_parent.index, ["a", "b"])
    c = ["a_n", "a_v", "a_s", "b_n", "b_v", "b_s"]
    assert df.columns.to_list() == c
    assert df.index.names == ["x", "y"]
    assert len(df) == 14


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("G4", 1),
        ("I4", 1),
        ("J4", 1),
        ("I7", 4),
        ("J7", 4),
        ("I17", 2),
        ("J17", 2),
        ("I13", 1),
        ("I14", 2),
        ("I15", 3),
        ("L13", 1),
        ("L14", 2),
        ("L15", 2),
        ("K4", norm.ppf(1 / 6)),
        ("N5", norm.ppf(2 / 6)),
        ("K6", norm.ppf(3 / 6)),
        ("N7", norm.ppf(4 / 6)),
        ("K8", norm.ppf(5 / 6)),
        ("K13", norm.ppf(1 / 4)),
        ("K14", norm.ppf(2 / 4)),
        ("K15", norm.ppf(3 / 4)),
        ("N13", norm.ppf(1 / 3)),
        ("N14", norm.ppf(2 / 3)),
        ("N15", norm.ppf(2 / 3)),
        ("N16", norm.ppf(1 / 3)),
        ("K17", norm.ppf(2 / 3)),
    ],
)
def test_value(sf: DistFrame, cell: str, value: float):
    v = sf.sheet[cell].value
    assert v is not None
    np.testing.assert_allclose(v, value)
