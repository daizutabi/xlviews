from __future__ import annotations

from xlviews.testing.common import create_sheet
from xlviews.testing.sheet_frame.base import Index, MultiColumn, NoIndex

if __name__ == "__main__":
    sheet = create_sheet()
    fc = NoIndex(sheet, style=True)
    print(fc.sf.melt())

    print(fc.sf.agg())

    fc = Index(sheet, style=True)
    print(fc.sf.groupby("name").agg())

    fc = MultiColumn(sheet, style=True)
    print(fc.sf.melt())

    print(fc.sf.agg())
