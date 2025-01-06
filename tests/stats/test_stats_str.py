import numpy as np
import pytest

from xlviews.frame import SheetFrame
from xlviews.stats import StatsFrame


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    return StatsFrame(sf_parent, by="x", stats="count", table=True)


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
        ("C3:C4", ["a", "b"]),
        ("D3:D4", [None, None]),
        ("E3:E4", [None, None]),
        ("F3:F4", [9, 9]),
        ("G3:G4", [10, 10]),
        ("H3:H4", [7, 10]),
    ],
)
def test_value(sf: SheetFrame, cell, value):
    assert sf.sheet[cell].value == value
