from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from xlviews.testing.stats_frame import Parent

if TYPE_CHECKING:
    from xlwings import Sheet

    from xlviews.testing import FrameContainer


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Parent(sheet_module, 3, 3, table=True)


@pytest.fixture(scope="module")
def df(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf_parent(fc: FrameContainer):
    return fc.sf
