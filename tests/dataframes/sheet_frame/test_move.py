import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.dataframes.sheet_frame import SheetFrame
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture
def sf(sheet: Sheet):
    df = DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})
    return SheetFrame(2, 3, data=df, sheet=sheet)


def test_range_original(sf: SheetFrame):
    assert sf.expand().get_address() == "$C$2:$E$6"


def test_move_down(sf: SheetFrame):
    sf.cell.offset(-1).value = "=$A$1"
    sf.move(20, "down")
    assert sf.cell.get_address() == "$C$23"
    assert sf.expand().get_address() == "$C$23:$E$27"
    assert sf.row == 23
    assert sf.column == 3
    v = [[None, "a", "b"], [0, 1, 5], [1, 2, 6], [2, 3, 7], [3, 4, 8]]
    assert sf.cell.expand().options(ndim=2).value == v
    assert sf.sheet.range("$C$2:$E$6").value == [[None] * 3] * 5


def test_move_right(sf: SheetFrame):
    sf.move(10, "right", width=3)
    assert sf.cell.get_address() == "$M$2"
    assert sf.expand().get_address() == "$M$2:$O$6"
    v = [[None, "a", "b"], [0, 1, 5], [1, 2, 6], [2, 3, 7], [3, 4, 8]]
    assert sf.cell.expand().options(ndim=2).value == v
    assert sf.sheet.range("$C$2:$E$6").value == [[None] * 3] * 5
    assert sf.sheet.range("C2").column_width == 3


def test_move_error(sf: SheetFrame):
    with pytest.raises(ValueError, match="direction must be 'down' or 'right'"):
        sf.move(10, "up")
