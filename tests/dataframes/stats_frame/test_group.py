from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from xlviews.dataframes.groupby import GroupBy
from xlviews.dataframes.stats_frame import get_by, get_frame
from xlviews.testing import is_app_available

if TYPE_CHECKING:
    from xlviews.dataframes.sheet_frame import SheetFrame

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def gr(sf_parent: SheetFrame):
    return GroupBy(sf_parent, ["x", "y"])


@pytest.mark.parametrize(
    ("funcs", "shape"),
    [(["mean"], (4, 3)), (["min", "max"], (8, 3))],
)
def test_get_frame_shape(gr: GroupBy, funcs: list[str], shape: tuple[int, int]):

    assert get_frame(gr, funcs).shape == shape


def test_get_frame_index_list(gr: GroupBy):
    df = get_frame(gr, ["mean"])
    assert df.index.names == ["func", "x", "y", "z"]


def test_get_frame_offset(gr: GroupBy):
    df = get_frame(gr, ["mean"]).reset_index()
    assert df["x"].iloc[0] == "=$C$4"
    assert df["y"].iloc[-1] == "=$D$16"
    assert df["a"].iloc[0] == "=AGGREGATE(1,7,$F$4:$F$7,$F$20:$F$23)"
    assert df["c"].iloc[-1] == "=AGGREGATE(1,7,$H$16:$H$19)"


def test_get_by_none(sf_parent: SheetFrame):
    assert get_by(sf_parent, None) == ["x", "y", "z"]
