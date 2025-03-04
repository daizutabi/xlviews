import numpy as np
import pytest
from scipy.stats import norm
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(params=["norm", "weibull"])
def dist(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture
def value(sheet_module: Sheet, dist: str):
    from xlviews.dataframes.dist_frame import set_formula, sigma_value

    rng = Range((1, 1), (10, 1), sheet_module)
    rng.value = [[x] for x in range(1, 11)]
    cell = rng[0]
    sigma = sigma_value(cell, 10, dist)
    set_formula(cell.offset(0, 1), 10, sigma)
    return Range((1, 2), (10, 2), sheet_module).value


@pytest.fixture
def expected(dist: str):
    if dist == "norm":
        return norm.ppf(np.arange(1, 11) / 11)

    return np.log(-np.log(1 - np.arange(1, 11) / 11))


def test_sigma_value(value, expected):
    np.testing.assert_allclose(value, expected)


def test_sigma_value_error(sheet_module: Sheet):
    from xlviews.dataframes.dist_frame import sigma_value

    with pytest.raises(ValueError, match="unknown distribution"):
        sigma_value(Range((1, 1), (1, 1), sheet_module), 10, "unknown")
