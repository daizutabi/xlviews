from itertools import product

import numpy as np
import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.heat_frame import HeatFrame
from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module")
def sf(sheet_module: Sheet):
    values = list(product(range(1, 5), range(1, 7)))
    df = DataFrame(values, columns=["x", "y"])
    df["v"] = list(range(len(df)))
    df = df[(df["x"] + df["y"]) % 4 != 0]
    df = df.set_index(["x", "y"])

    sf = SheetFrame(2, 2, data=df, index=True, sheet=sheet_module)
    data = sf.get_address(["v"], formula=True)

    return HeatFrame(2, 6, data=data, x="x", y="y", value="v", sheet=sheet_module)


def test_index(sf: HeatFrame):
    assert sf.sheet.range("F3:F8").value == [1, 2, 3, 4, 5, 6]


def test_columns(sf: HeatFrame):
    assert sf.sheet.range("G2:J2").value == [1, 2, 3, 4]


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, [0, 6, None, 18]),
        (4, [1, None, 13, 19]),
        (5, [None, 8, 14, 20]),
        (6, [3, 9, 15, None]),
        (7, [4, 10, None, 22]),
        (8, [5, None, 17, 23]),
    ],
)
def test_values(sf: HeatFrame, i: int, value: int):
    assert sf.sheet.range(f"G{i}:J{i}").value == value


def test_vmin(sf: HeatFrame):
    assert sf.vmin.get_address() == "$L$8"


def test_vmax(sf: HeatFrame):
    assert sf.vmax.get_address() == "$L$3"


def test_label(sf: HeatFrame):
    assert sf.label.get_address() == "$L$2"


def test_label_value(sf: HeatFrame):
    assert sf.label.value == "v"


@pytest.mark.parametrize(
    ("i", "value"),
    [
        (3, 23),
        (4, 23 * 4 / 5),
        (5, 23 * 3 / 5),
        (6, 23 * 2 / 5),
        (7, 23 / 5),
        (8, 0),
    ],
)
def test_colorbar(sf: HeatFrame, i: int, value: int):
    v = sf.sheet.range(f"L{i}").value
    assert isinstance(v, float)
    np.testing.assert_allclose(v, value)


if __name__ == "__main__":
    from xlviews.testing import create_sheet

    sheet = create_sheet()

    values = list(product(range(1, 5), range(1, 7)))
    df = DataFrame(values, columns=["x", "y"])
    df["v"] = list(range(len(df)))
    df = df[(df["x"] + df["y"]) % 4 != 0]
    df = df.set_index(["x", "y"])
    sf = SheetFrame(2, 2, data=df, index=True)  # type: ignore

    data = sf.get_address(["v"], formula=True)
    hf = HeatFrame(2, 6, data=data, x="x", y="y", value="v")

    hf.range(index=False)
