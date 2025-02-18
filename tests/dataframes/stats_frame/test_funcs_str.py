import pytest

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.dataframes.stats_frame import StatsFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    return StatsFrame(sf_parent, "count", by="x")


def test_len(sf: SheetFrame):
    assert len(sf) == 2


def test_index_names(sf: SheetFrame):
    assert sf.index.names == ["func", "x", "y", "z"]


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("C3:C5", ["x", "a", "b"]),
        ("D3:D5", ["y", None, None]),
        ("E3:E5", ["z", None, None]),
        ("F3:F5", ["a", 10, 8]),
        ("G3:G5", ["b", 12, 8]),
        ("H3:H5", ["c", 10, 7]),
    ],
)
def test_value(sf: SheetFrame, cell, value):
    assert sf.sheet[cell].value == value
