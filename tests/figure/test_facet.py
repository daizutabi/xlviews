import pytest
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.figure.plot import Plot
from xlviews.testing import is_app_available
from xlviews.testing.figure.facet import facet
from xlviews.testing.sheet_frame.pivot import Pivot

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
def plots(sf: SheetFrame):
    return [plot for _, plot in facet(sf)]


def test_plots_len(plots: list[Plot]):
    assert len(plots) == 5


@pytest.mark.parametrize(
    ("i", "title"),
    [(0, "1_1"), (1, "2_1"), (2, "1_2"), (3, "2_2"), (4, "2_3")],
)
def test_title(plots: list[Plot], i: int, title: str):
    assert plots[i].axes.title == title


def test_iterrows_none():
    from pandas import Index

    from xlviews.figure.plot import iterrows

    assert list(iterrows(Index([]), None)) == [{}]
