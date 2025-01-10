import pytest
from pandas import DataFrame
from xlwings import Sheet

from xlviews.frame import SheetFrame


@pytest.fixture(scope="module")
def df():
    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    return df.set_index(["x", "y"])


@pytest.fixture(scope="module")
def sf(df: DataFrame, sheet_module: Sheet):
    return SheetFrame(sheet_module, 3, 2, data=df, style=False)


if __name__ == "__main__":
    import xlwings as xw
    from pandas import DataFrame

    from xlviews.common import quit_apps
    from xlviews.dist import DistFrame
    from xlviews.frame import SheetFrame

    quit_apps()
    book = xw.Book()
    sheet_module = book.sheets.add()

    df = DataFrame(
        {
            "x": [1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2],
            "y": [3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 3, 3, 4, 4],
            "a": [5, 4, 3, 2, 1, 4, 3, 2, 1, 3, 2, 1, 2, 1],
            "b": [1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3, 1, 2],
        },
    )
    df = df.set_index(["x", "y"])

    sf = SheetFrame(sheet_module, 3, 2, data=df, style=False)

    sfd = DistFrame(sf, "a", by=["x", "y"])
    sfd.columns
