from __future__ import annotations

import pandas as pd
from pandas import DataFrame
from xlwings.constants import Direction

from xlviews.config import rcParams
from xlviews.core.formula import AGG_FUNCS
from xlviews.core.range_collection import RangeCollection
from xlviews.style import set_font, set_number_format
from xlviews.utils import iter_columns, suspend_screen_updates

from .groupby import GroupBy
from .sheet_frame import SheetFrame


class StatsFrame(SheetFrame):
    @suspend_screen_updates
    def __init__(
        self,
        parent: SheetFrame,
        funcs: str | list[str] | None = None,
        by: str | list[str] | None = None,
        *,
        default: str = "median",
        func_column_name: str = "func",
        auto_filter: bool = True,
    ) -> None:
        """Create a StatsFrame.

        Args:
            parent (SheetFrame): The `SheetFrame` to be aggregated.
            funcs (str, list of str, optional): The aggregation
                functions to be used. The following functions are supported:
                    'count', 'sum', 'min', 'max', 'mean', 'median', 'std', 'soa%'
                None to use the default functions.
            by (str, list of str, optional): The column names to be grouped by.
            default (str, optional): The default function to be displayed.
            func_column_name (str, optional): The name of the function column.
            auto_filter (bool, optional): Whether to automatically filter the data.
        """
        funcs = get_func(funcs)
        by = get_by(parent, by)
        offset = get_length(parent, by, funcs) + 2

        # Store the position of the parent SheetFrame before moving down.
        row = parent.row
        column = parent.column
        if isinstance(funcs, list):
            column -= 1

        move_down(parent, offset)

        gp = GroupBy(parent, by)
        data = get_frame(gp, funcs, func_column_name)

        super().__init__(row, column, data, parent.sheet)

        self.as_table(autofit=False, const_header=True)
        self.style()

        if isinstance(funcs, list):
            set_style(self, parent, func_column_name)

        self.alignment("left")

        if self.table and auto_filter and isinstance(funcs, list) and len(funcs) > 1:
            func = default if default in funcs else funcs[0]
            self.table.auto_filter(func_column_name, func)


def get_func(func: str | list[str] | None) -> list[str]:
    if func is None:
        func = list(AGG_FUNCS.keys())
        func.remove("sum")
        func.remove("std")
        return func

    return [func] if isinstance(func, str) else func


def get_by(sf: SheetFrame, by: str | list[str] | None) -> list[str]:
    if not by:
        return sf.index.names

    return list(iter_columns(sf.index.names, by))


def get_length(sf: SheetFrame, by: list[str], funcs: list | dict) -> int:
    n = 1 if isinstance(funcs, dict) else len(funcs)

    if not by:
        return n

    return len(sf.index.to_frame()[by].drop_duplicates()) * n


def get_frame(
    group: GroupBy,
    funcs: list[str],
    func_column_name: str = "func",
) -> DataFrame:
    df = group.agg(funcs, group.sf.columns.to_list(), formula=True, as_address=True)
    df = df.stack(level=1, future_stack=True)  # noqa: PD013

    index = df.index.to_frame()
    index = index.rename(columns={index.columns[-1]: func_column_name})

    for c in group.sf.index.names:
        if c not in group.by:
            index[c] = ""

    df = pd.concat([index, df], axis=1)
    return df.set_index([func_column_name, *group.sf.index.names])


def has_header(sf: SheetFrame) -> bool:
    start = sf.cell.offset(-1)
    end = start.offset(0, sf.index.nlevels)
    value = sf.sheet.range(start, end).options(ndim=1).value

    if not isinstance(value, list):
        raise NotImplementedError

    return any(value)


def move_down(sf: SheetFrame, length: int) -> int:
    start = sf.row - 1
    end = sf.row + length - 2

    if has_header(sf):
        end += 1

    rows = sf.sheet.api.Rows(f"{start}:{end}")
    rows.Insert(Shift=Direction.xlDown)
    sf.cell = sf.cell.offset()  # update cell
    return end - start + 1


def set_style(sf: SheetFrame, parent: SheetFrame, func_column_name: str) -> None:
    idx = [sf.column + i for i in range(sf.index.nlevels + len(sf.columns))]

    columns = (*parent.index.names, *parent.columns)
    formats = [None, *[parent.get_number_format(column) for column in columns]]

    for (func,), rows in sf.groupby(func_column_name).items():
        for col, fmt in zip(idx, formats, strict=True):
            rc = RangeCollection(rows, col, sf.sheet)

            if func in ["median", "min", "mean", "max", "std", "sum"] and fmt:
                set_number_format(rc, fmt)

            color = rcParams.get(f"stats.{func}.color")
            italic = rcParams.get(f"stats.{func}.italic")
            set_font(rc, color=color, italic=italic)

            if func == "soa" and col != idx[0]:
                set_number_format(rc, "0.0%")

    rng = sf.get_range(func_column_name)
    set_font(rng, italic=True)
