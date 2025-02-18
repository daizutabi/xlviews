import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.dist_frame import DistFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available
from xlviews.testing.dist_frame import Parent

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.mark.parametrize(
    ("columns", "values"),
    [
        (None, ["a_n", "a_v", "a_s", "b_n", "b_v", "b_s"]),
        ("a", ["a_n", "a_v", "a_s"]),
        ("b", ["b_n", "b_v", "b_s"]),
    ],
)
def test_columns(columns, values, sheet: Sheet):
    fc = Parent(sheet, 3, 2)
    sf = DistFrame(fc.sf, columns, by=["x", "y"])
    assert sf.columns.to_list() == values


def test_index_str(sheet: Sheet):
    fc = Parent(sheet, 3, 2)
    sf = DistFrame(fc.sf, by="x")
    assert sf.index.names == ["x"]


def test_index_none(sheet: Sheet):
    fc = Parent(sheet, 3, 2)
    sf = DistFrame(fc.sf)
    assert sf.index.names == [None]
    assert sf.index.to_list() == list(range(14))


def test_group_error(sheet: Sheet):
    df = DataFrame(
        {
            "x": [1, 1, 1, 2, 2, 2, 1, 1, 1, 2, 2, 2, 2, 2],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    df = df.set_index(["x", "y"])
    sf = SheetFrame(3, 2, data=df, sheet=sheet)
    with pytest.raises(ValueError, match="group must be continuous"):
        sf = DistFrame(sf, None, by=["x", "y"])
