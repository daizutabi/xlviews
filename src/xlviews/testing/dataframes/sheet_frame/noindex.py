from __future__ import annotations

from pandas import DataFrame

from xlviews.testing import create_sheet, create_sheet_frame


def create_data_frame() -> DataFrame:
    values = {"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]}
    return DataFrame(values)


if __name__ == "__main__":
    sheet = create_sheet()
    df = create_data_frame()
    sf = create_sheet_frame(df, sheet, 2, 3)
