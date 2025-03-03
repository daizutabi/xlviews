import pytest
from xlwings import Sheet

from xlviews.core.range import Range
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def rng(sheet_module: Sheet):
    rng = Range((2, 2), (3, 3), sheet_module)
    rng.value = [[1, 2], [3, 4]]
    return rng


def test_colorbar_vertical(rng: Range, sheet_module: Sheet):
    from xlviews.dataframes.colorbar import Colorbar

    cb = Colorbar(2, 5, 6, sheet=sheet_module)
    cb.set(vmin=rng, vmax=rng, label="T", autofit=True)
    cb.set_adjacent_column_width(1)
    cb.apply(rng)

    assert sheet_module.range((1, 5)).value == "T"
    assert sheet_module.range((2, 5)).value == 4
    assert sheet_module.range((7, 5)).value == 1
    assert sheet_module.range((1, 6)).column_width == 1


def test_colorbar_horizontal(rng: Range, sheet_module: Sheet):
    from xlviews.dataframes.colorbar import Colorbar

    cb = Colorbar(2, 7, 10, orientation="horizontal", sheet=sheet_module)
    cb.set(vmin=rng, vmax=rng, label="T", autofit=True)
    cb.set_adjacent_column_width(1)

    assert sheet_module.range((2, 17)).value == "T"
    assert sheet_module.range((2, 16)).value == 4
    assert sheet_module.range((2, 7)).value == 1
    assert sheet_module.range((1, 18)).column_width == 1
