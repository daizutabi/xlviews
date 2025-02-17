import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(
    scope="module",
    params=[(10, 10), (20, 10), (50, 10), (500, 500)],
    ids=lambda x: str(x),
)
def shape(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def rng_impl(shape: tuple[int, int], sheet_module: Sheet):
    nrows, ncolumns = shape
    return sheet_module.range((1, 1), (nrows, ncolumns))


@pytest.fixture(scope="module")
def rng(shape: tuple[int, int], sheet_module: Sheet):
    nrows, ncolumns = shape
    return Range((1, 1), (nrows, ncolumns), sheet=sheet_module)


@pytest.mark.slow
def test_address_impl(benchmark, rng_impl: RangeImpl, rng: Range, shape):
    x = benchmark(lambda: [r.get_address() for r in rng_impl])
    assert len(x) == shape[0] * shape[1]
    assert x == list(rng.iter_addresses())


def test_address(benchmark, rng: Range, shape):
    x = benchmark(lambda: list(rng.iter_addresses()))
    assert len(x) == shape[0] * shape[1]


def test_address_frame(benchmark, rng: Range, shape):
    x = benchmark(lambda: rng.frame.get_address())
    assert x.shape == (shape[0], shape[1])


def test_address_iter_frame(benchmark, rng: Range, shape):
    def f():
        x = np.array(list(rng.iter_addresses())).reshape(shape[0], shape[1])
        return DataFrame(x)

    x = benchmark(f)
    assert x.shape == (shape[0], shape[1])
