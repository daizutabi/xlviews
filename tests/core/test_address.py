import pytest

from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize("index", range(1, 1000, 50))
def test_column_name(index: int):
    from xlviews.core.address import column_name_to_index, index_to_column_name

    assert column_name_to_index(index_to_column_name(index)) == index
