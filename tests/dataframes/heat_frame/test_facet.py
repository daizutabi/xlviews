import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.testing import is_app_available
from xlviews.testing.heat_frame.facet import FacetParent

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def fc_parent(sheet_module: Sheet):
    return FacetParent(sheet_module)


@pytest.fixture(scope="module")
def df(fc_parent: FacetParent):
    return fc_parent.df
