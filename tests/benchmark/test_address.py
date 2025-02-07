import pytest
from xlwings import Range, Sheet

from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(
    scope="module",
    params=[(10, 10), (30, 10), (100, 20)],
    ids=lambda x: str(x),
)
def shape(request: pytest.FixtureRequest):
    return request.param


def test_get_addresses(benchmark, sheet: Sheet, shape: tuple[int, int]):
    nrows, ncolumns = shape
    rng = sheet.range((1, 1), (nrows, ncolumns))
    x = benchmark(lambda rng: [x.get_address() for x in rng], rng)
    assert len(x) == nrows * ncolumns


def test_iter_addresses(benchmark, sheet: Sheet, shape: tuple[int, int]):
    from xlviews.range.address import iter_addresses

    nrows, ncolumns = shape
    rng = sheet.range((1, 1), (nrows, ncolumns))
    x = benchmark(lambda x: list(iter_addresses(x, cellwise=True)), rng)
    assert len(x) == nrows * ncolumns
