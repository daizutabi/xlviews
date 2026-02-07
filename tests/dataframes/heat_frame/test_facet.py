from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.facet import facet
from xlviews.testing.sheet_frame.pivot import Pivot

if TYPE_CHECKING:
    from xlwings import Sheet

    from xlviews.dataframes.heat_frame import HeatFrame
    from xlviews.dataframes.sheet_frame import SheetFrame

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Pivot(sheet_module)


@pytest.fixture(scope="module")
def df(fc: Pivot):
    return fc.df


@pytest.fixture(scope="module")
def sf(fc: Pivot):
    return fc.sf


@pytest.fixture(scope="module")
def hfs(sf: SheetFrame):
    return [hf for _, hf in facet(sf)]


@pytest.mark.parametrize(
    ("i", "v"),
    [(0, 1411), (1, 2511), (2, 1711), (3, 2811), (4, None), (5, 3111)],
)
def test_value(hfs: list[HeatFrame], i: int, v: int | None):
    assert hfs[i].cell.offset(1, 1).value == v
