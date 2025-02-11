import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.groupby import groupby
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


def test_groupby(sheet: Sheet):
    df = DataFrame({"a": [1, 1, 1, 2, 2, 1, 1], "b": [1, 2, 3, 4, 5, 6, 7]})
    sf = SheetFrame(2, 2, data=df, index=False, sheet=sheet)

    g = groupby(sf, "a")
    assert len(g) == 2
    assert g[(1,)] == [(3, 5), (8, 9)]
    assert g[(2,)] == [(6, 7)]

    assert len(groupby(sf, ["a", "b"])) == 7

    g = groupby(sf, "::b")
    assert len(g) == 2
    assert g[(1,)] == [(3, 5), (8, 9)]
    assert g[(2,)] == [(6, 7)]

    assert len(groupby(sf, ":b")) == 7
