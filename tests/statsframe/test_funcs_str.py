import pytest

from xlviews.sheetframe import SheetFrame
from xlviews.statsframe import StatsFrame


@pytest.fixture(scope="module", params=[True, False])
def sf(sf_parent: SheetFrame, request: pytest.FixtureRequest):
    return StatsFrame(sf_parent, "count", by="x", table=request.param)


def test_len(sf: SheetFrame):
    assert len(sf) == 2


def test_columns(sf: SheetFrame):
    assert sf.columns == ["func", "x", "y", "z", "a", "b", "c"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["func", "x", "y", "z"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b", "c"]


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
