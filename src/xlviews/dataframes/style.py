"""Set styles for SheetFrame."""

from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING

import pywintypes
from xlwings.constants import TableStyleElementType

from xlviews.config import rcParams
from xlviews.style import (
    EVEN_COLOR,
    ODD_COLOR,
    hide_succession,
    hide_unique,
    set_alignment,
    set_banding,
    set_border,
    set_fill,
    set_font,
)
from xlviews.utils import iter_group_locs, suspend_screen_updates

if TYPE_CHECKING:
    from pandas import Index
    from xlwings import Range, Sheet

    from xlviews.colors import Color

    from .heat_frame import HeatFrame
    from .sheet_frame import SheetFrame
    from .table import Table


def _set_style(
    start: Range,
    end: Range | None,
    name: str,
    *,
    border: bool = True,
    fill: bool = True,
    font: bool = True,
    font_size: int | None = None,
) -> None:
    rng = start.sheet.range(start, end)

    if border:
        edge_color = rcParams["frame.border.color"]
        inside_color = rcParams["frame.border.inside.color"]
        set_border(rng, edge_color=edge_color, inside_color=inside_color)

    if fill:
        set_fill(rng, color=rcParams[f"frame.{name}.fill.color"])

    if font:
        color = rcParams[f"frame.{name}.font.color"]
        bold = rcParams[f"frame.{name}.font.bold"]
        size = font_size or rcParams["frame.font.size"]
        set_font(rng, color=color, bold=bold, size=size)


@suspend_screen_updates
def set_frame_style(
    sf: SheetFrame,
    *,
    border: bool = True,
    fill: bool = True,
    font: bool = True,
    font_size: int | None = None,
    alignment: str | None = "center",
    banding: bool = False,
    succession: bool = False,
) -> None:
    """Set style of SheetFrame.

    Args:
        sf: The SheetFrame object.
        border: Whether to draw the border.
        font: Whether to specify the font.
        fill: Whether to fill the frame.
        font_size: The font size to specify directly.
        alignment: The alignment of the frame.
        banding: Whether to draw the banding.
        succession: Whether to hide the succession of the index.
    """
    cell = sf.cell
    sheet = sf.sheet

    set_style = partial(
        _set_style,
        border=border,
        fill=fill,
        font=font,
        font_size=font_size,
    )

    index_nlevels = sf.index.nlevels
    columns_nlevels = sf.columns.nlevels
    length = len(sf)

    end = cell.offset(columns_nlevels - 1, index_nlevels - 1)
    if columns_nlevels > 1 and index_nlevels == 1:
        set_style(cell, end, "columns.name")
    else:
        set_style(cell, end, "index.name")

    start = cell.offset(columns_nlevels, 0)
    end = cell.offset(columns_nlevels + length - 1, index_nlevels - 1)
    set_style(start, end, "index")

    if succession:
        rng = sheet.range(start.offset(1, 0), end)
        hide_succession(rng)

        start = cell.offset(columns_nlevels - 1, 0)
        end = cell.offset(columns_nlevels - 1, index_nlevels - 1)
        rng = sheet.range(start, end)
        hide_unique(rng, length)

    width = len(sf.columns)

    start = cell.offset(0, index_nlevels)
    end = cell.offset(columns_nlevels - 1, index_nlevels + width - 1)
    set_style(start, end, "columns")

    start = cell.offset(columns_nlevels, index_nlevels)
    end = cell.offset(columns_nlevels + length - 1, index_nlevels + width - 1)
    set_style(start, end, "values")

    rng = sheet.range(start, end)

    if banding:
        set_banding(rng)

    rng = sheet.range(cell, end)

    if border:
        ew = rcParams["frame.border.weight"]
        ec = rcParams["frame.border.color"]
        set_border(rng, edge_weight=ew, inside_weight=0, edge_color=ec)

    if alignment:
        set_alignment(rng, alignment)


def set_wide_column_style(sf: SheetFrame) -> None:
    edge_color = rcParams["frame.border.color"]
    edge_weight = rcParams["frame.wide-columns.border.weight"]

    columns = list(sf.columns.wide_index)
    for column in columns:
        loc = sf.columns.get_loc(column, sf.column + sf.index.nlevels)
        if not isinstance(loc, tuple):
            raise NotImplementedError

        rng = sf.sheet.range((sf.row, loc[0]), (sf.row, loc[1]))

        er = edge_weight if column == columns[-1] else 2
        edge_weight_tuple = (2, er, 2, 2)
        set_border(rng, edge_weight_tuple, inside_weight=1, edge_color=edge_color)
        _set_style(rng, None, "wide-columns", border=False)

        rng = sf.sheet.range((sf.row - 1, loc[0]), (sf.row - 1, loc[1]))

        el = edge_weight if column == columns[0] else 2
        edge_weight_tuple = (el, edge_weight, edge_weight, 2)
        set_border(rng, edge_weight_tuple, inside_weight=0, edge_color=edge_color)
        _set_style(rng, None, "wide-columns.name", border=False)


def set_table_style(
    table: Table,
    even_color: Color = EVEN_COLOR,
    odd_color: Color = ODD_COLOR,
) -> None:
    book = table.sheet.book.api

    try:
        style = book.TableStyles("xlviews")
    except pywintypes.com_error:
        style = book.TableStyles.Add("xlviews")
        odd_type = TableStyleElementType.xlRowStripe1
        style.TableStyleElements(odd_type).Interior.Color = odd_color
        even_type = TableStyleElementType.xlRowStripe2
        style.TableStyleElements(even_type).Interior.Color = even_color

    table.api.TableStyle = style


@suspend_screen_updates
def set_heat_frame_style(
    sf: HeatFrame,
    *,
    border: bool = True,
    fill: bool = True,
    font: bool = True,
    font_size: int | None = None,
    alignment: str | None = "center",
) -> None:
    """Set style of SheetFrame.

    Args:
        sf: The SheetFrame object.
        border: Whether to draw the border.
        fill: Whether to fill the frame.
        font: Whether to specify the font.
        font_size: The font size to specify directly.
        alignment: The alignment of the frame.
    """
    cell = sf.cell
    sheet = sf.sheet

    set_style = partial(
        _set_style,
        border=border,
        fill=fill,
        font=font,
        font_size=font_size,
    )

    index_nlevels = sf.index.nlevels
    columns_nlevels = sf.columns.nlevels
    length = len(sf)

    start = cell.offset(columns_nlevels, 0)
    end = cell.offset(columns_nlevels + length - 1, index_nlevels - 1)
    set_style(start, end, "index")

    width = len(sf.columns)

    start = cell.offset(columns_nlevels - 1, index_nlevels)
    end = cell.offset(columns_nlevels - 1, index_nlevels + width - 1)
    set_style(start, end, "index")

    start = cell.offset(columns_nlevels, index_nlevels)
    end = cell.offset(columns_nlevels + length - 1, index_nlevels + width - 1)
    set_style(start, end, "values")

    rng = sheet.range(cell, end)

    if alignment:
        set_alignment(rng, alignment)

    _merge_index(sf.columns, sf.row, sf.column, 1, sf.sheet)
    _merge_index(sf.index, sf.row, sf.column, 0, sf.sheet)
    _set_heat_border(sf)


def _merge_index(index: Index, row: int, column: int, axis: int, sheet: Sheet) -> None:
    for start, end in iter_group_locs(index):
        if start == end:
            continue
        if axis == 0:
            sheet.range((row + start + 1, column), (row + end + 1, column)).merge()
        else:
            sheet.range((row, column + start + 1), (row, column + end + 1)).merge()


def _set_heat_border(sf: HeatFrame) -> None:
    r = sf.row + 1
    c = sf.column + 1

    ec = rcParams["heat.border.color"]

    for row in iter_group_locs(sf.index, offset=r):
        for col in iter_group_locs(sf.columns, offset=c):
            if row[0] == row[1] and col[0] == col[1]:
                continue

            rng = sf.sheet.range((row[0], col[0]), (row[1], col[1]))
            set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)
