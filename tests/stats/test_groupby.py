import numpy as np
import pytest

from xlviews.formula import NONCONST_VALUE
from xlviews.frame import SheetFrame
from xlviews.stats import StatsFrame


@pytest.fixture(scope="module")
def sf(sf_parent: SheetFrame):
    stats = {"a": "max", "b": "std", "c": "mean"}
    return StatsFrame(sf_parent, by=":y", stats=stats, table=True)


def test_columns(sf: SheetFrame):
    assert sf.columns == ["x", "y", "z", "a", "b", "c"]


def test_value_columns(sf: SheetFrame):
    assert sf.value_columns == ["a", "b", "c"]


def test_index_columns(sf: SheetFrame):
    assert sf.index_columns == ["x", "y", "z"]


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("C3:C6", ["a", "a", "b", "b"]),
        ("D3:D6", ["c", "d", "c", "d"]),
        ("E3:E6", [None, None, None, None]),
        ("F3:F6", [5, 9, 15, 18]),
        ("H3:H6", [24.8, 35, 5, 15]),
    ],
)
def test_value(sf: SheetFrame, cell, value):
    assert sf.sheet[cell].value == value


@pytest.mark.parametrize(
    ("cell", "value"),
    [
        ("G3", np.std(range(6))),
        ("G4", np.std(range(4))),
        ("G5", np.std(range(0, 18, 3))),
        ("G6", np.std(range(18, 30, 3))),
    ],
)
def test_value_std(sf: SheetFrame, cell, value):
    v = sf.sheet[cell].value
    assert v is not None
    np.testing.assert_allclose(v, value)


@pytest.mark.parametrize("cell", ["C1", "D1"])
def test_const_header_nonconst(sf: SheetFrame, cell):
    assert sf.sheet.range(cell).value == NONCONST_VALUE


@pytest.mark.parametrize("cell", ["E1", "F1", "G1", "H1"])
def test_const_header_error(sf: SheetFrame, cell):
    assert sf.sheet.range(cell).value is None
