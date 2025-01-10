import numpy as np
import pytest
from scipy.stats import norm

from xlviews.distframe import DistFrame
from xlviews.sheetframe import SheetFrame


def test_init_data(sf: SheetFrame):
    from xlviews.distframe import get_init_data

    df = get_init_data(sf, ["a", "b"], ["x", "y"])
    c = ["a_n", "a_v", "a_s", "b_n", "b_v", "b_s"]
    assert df.columns.to_list() == c
    assert df.index.names == ["x", "y"]
    assert len(df) == 14


def test_dist_func_str():
    from xlviews.distframe import get_dist_func

    df = get_dist_func("norm", ["a", "b"])
    assert df == {"a": "norm", "b": "norm"}


def test_dist_func_dict():
    from xlviews.distframe import get_dist_func

    df = get_dist_func({"a": "none"}, ["a", "b"])
    assert df == {"a": "none", "b": "norm"}


@pytest.fixture(scope="module")
def sfd(sf: SheetFrame):
    from xlviews.distframe import DistFrame

    return DistFrame(sf, ["a", "b"], by=["x", "y"])


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
        ("K4", norm.ppf(1 / 6)),
        ("N5", norm.ppf(2 / 6)),
        ("K6", norm.ppf(3 / 6)),
        ("N7", norm.ppf(4 / 6)),
        ("K8", norm.ppf(5 / 6)),
        ("N16", norm.ppf(1 / 3)),
        ("K17", norm.ppf(2 / 3)),
    ],
)
def test_distframe(sfd: DistFrame, cell: str, value: float):
    v = sfd.sheet[cell].value
    assert v is not None
    np.testing.assert_allclose(v, value)
