import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.dist_frame import DistFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import FrameContainer
from xlviews.testing.dist_frame import Parent


@pytest.fixture(scope="module")
def fc(sheet_module: Sheet):
    return Parent(sheet_module, 3, 2, style=True)


@pytest.fixture(scope="module")
def df_parent(fc: FrameContainer):
    return fc.df


@pytest.fixture(scope="module")
def sf_parent(fc: FrameContainer):
    return fc.sf


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    return DistFrame(sf_parent, ["a", "b"], by=["x", "y"])


@pytest.fixture(scope="module")
def df(sf: DistFrame):
    rng = sf.expand().impl
    return rng.options(DataFrame, index=sf.index.nlevels).value
