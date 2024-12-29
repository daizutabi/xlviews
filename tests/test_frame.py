import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df1():
    return DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})


def test_init_df1(df1: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df1, style=False)
    assert sf.row == 2
    assert sf.column == 3
    assert sf.sheet.name == sheet.name
    assert sf.columns == [None, "a", "b"]
    assert len(sf) == 3
    assert "a" in sf
    assert "x" not in sf


def test_init_df1_index_false(df1: DataFrame, sheet: Sheet):
    sf = SheetFrame(sheet, 2, 3, data=df1, index=False, style=False)
    assert sf.columns == ["a", "b"]
