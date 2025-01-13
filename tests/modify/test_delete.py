import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.sheetframe import SheetFrame


@pytest.fixture
def sfs(sheet: Sheet):
    df0 = DataFrame({"a": [1, 1], "b": [1, 1]})
    df0 = df0.set_index("a")
    df1 = DataFrame({"c": [2, 2], "d": [2, 2]})
    df1 = df1.set_index("c")
    df2 = DataFrame({"e": [3, 3], "f": [3, 3]})
    df2 = df2.set_index("e")

    sf0 = SheetFrame(2, 3, data=df0, style=False, sheet=sheet)
    sf1 = SheetFrame(7, 3, data=df1, style=False, sheet=sheet)
    sf2 = SheetFrame(2, 7, data=df2, style=False, sheet=sheet)

    return [sf0, sf1, sf2]


def test_sfs_value(sfs: list[SheetFrame]):
    v = sfs[0].sheet.range("$C$2:$I$10").value
    assert isinstance(v, list)
    assert len(v) == 9
    assert v[0] == ["a", "b", None, None, "e", "f", None]
    assert v[1] == [1, 1, None, None, 3, 3, None]
    assert v[2] == [1, 1, None, None, 3, 3, None]
    assert v[3] == [None] * 7
    assert v[4] == [None] * 7
    assert v[5] == ["c", "d", None, None, None, None, None]
    assert v[6] == [2, 2, None, None, None, None, None]
    assert v[7] == [2, 2, None, None, None, None, None]
    assert v[8] == [None] * 7


def test_delete_up(sfs: list[SheetFrame]):
    sfs[0].delete("up")

    v = sfs[0].sheet.range("$C$2:$I$10").value
    assert isinstance(v, list)
    assert len(v) == 9
    assert v[0] == ["c", "d", None, None, "e", "f", None]
    assert v[1] == [2, 2, None, None, 3, 3, None]
    assert v[2] == [2, 2, None, None, 3, 3, None]
    assert v[3] == [None] * 7
    assert v[4] == [None] * 7
    assert v[5] == [None] * 7
    assert v[6] == [None] * 7
    assert v[7] == [None] * 7
    assert v[8] == [None] * 7


def test_delete_left(sfs: list[SheetFrame]):
    sfs[0].delete("left")

    v = sfs[0].sheet.range("$C$2:$I$10").value
    assert isinstance(v, list)
    assert len(v) == 9
    assert v[0] == ["e", "f", None, None, None, None, None]
    assert v[1] == [3, 3, None, None, None, None, None]
    assert v[2] == [3, 3, None, None, None, None, None]
    assert v[3] == [None] * 7
    assert v[4] == [None] * 7
    assert v[5] == ["c", "d", None, None, None, None, None]
    assert v[6] == [2, 2, None, None, None, None, None]
    assert v[7] == [2, 2, None, None, None, None, None]
    assert v[8] == [None] * 7


def test_delete_up_entire(sfs: list[SheetFrame]):
    sfs[0].delete("up", entire=True)

    v = sfs[0].sheet.range("$C$2:$I$10").value
    assert isinstance(v, list)
    assert len(v) == 9
    assert v[0] == ["c", "d", None, None, None, None, None]
    assert v[1] == [2, 2, None, None, None, None, None]
    assert v[2] == [2, 2, None, None, None, None, None]
    assert v[3] == [None] * 7
    assert v[4] == [None] * 7
    assert v[5] == [None] * 7
    assert v[6] == [None] * 7
    assert v[7] == [None] * 7
    assert v[8] == [None] * 7


def test_delete_left_entire(sfs: list[SheetFrame]):
    sfs[0].delete("left", entire=True)

    v = sfs[0].sheet.range("$C$2:$I$10").value
    assert isinstance(v, list)
    assert len(v) == 9
    assert v[0] == ["e", "f", None, None, None, None, None]
    assert v[1] == [3, 3, None, None, None, None, None]
    assert v[2] == [3, 3, None, None, None, None, None]
    assert v[3] == [None] * 7
    assert v[4] == [None] * 7
    assert v[5] == [None] * 7
    assert v[6] == [None] * 7
    assert v[7] == [None] * 7
    assert v[8] == [None] * 7


def test_delete_error(sfs: list[SheetFrame]):
    with pytest.raises(ValueError, match="direction must be 'up' or 'left'"):
        sfs[0].delete("right")
