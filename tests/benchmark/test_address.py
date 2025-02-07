import pytest
from xlwings import Range, Sheet

from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def get_addresses(rng: Range, **kwargs):
    return [r.get_address(**kwargs) for r in rng]


@pytest.mark.parametrize(("rows", "columns"), [(10, 10), (30, 10), (100, 10)])
def test_get_addresses(benchmark, sheet: Sheet, rows: int, columns: int):
    rng = sheet.range((1, 1), (rows, columns))
    assert rng[0].get_address() == "$A$1"
    x = benchmark(get_addresses, rng)
    assert x[0] == "$A$1"
