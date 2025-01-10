import numpy as np
import pytest

from xlviews.sheetframe import SheetFrame
from xlviews.statsframe import StatsFrame


@pytest.fixture(scope="module", params=[True, False])
def sf(sf_parent: SheetFrame, request: pytest.FixtureRequest):
    funcs = {"a": "max", "b": "std", "c": "mean"}
    return StatsFrame(sf_parent, funcs, by=":y", table=request.param)


def test_len(sf: SheetFrame):
    assert len(sf) == 4


def test_columns(sf: SheetFrame):
    assert sf.columns == ["x", "y", "z", "a", "b", "c"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y", "z"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b", "c"]


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("C3:C7", ["x", "a", "a", "b", "b"]),
        ("D3:D7", ["y", "c", "d", "c", "d"]),
        ("E3:E7", ["z", None, None, None, None]),
        ("F3:F7", ["a", 18, 7, 11, 15]),
        ("H3:H7", ["c", 18, (28 + 30 + 34) / 3, 38 / 3, 7]),
    ],
)
def test_value(sf: SheetFrame, cell, value):
    assert sf.sheet[cell].value == value


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("G4", np.std([0, 1, 2, 3, 18, 21, 24, 27])),
        ("G5", np.std([4, 5, 6, 7])),
        ("G6", np.std([8, 9, 0, 3])),
        ("G7", np.std([6, 9, 12, 15])),
    ],
)
def test_value_std(sf: SheetFrame, cell, value):
    v = sf.sheet[cell].value
    assert v is not None
    np.testing.assert_allclose(v, value)
