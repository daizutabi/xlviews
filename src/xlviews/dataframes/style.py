"""Set styles for SheetFrame."""

from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING

import pywintypes
from xlwings.constants import TableStyleElementType

from xlviews.colors import rgb
from xlviews.config import rcParams
from xlviews.core.style import (
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

    from .heat_frame import HeatFrame
    from .sheet_frame import SheetFrame
    from .table import Table


def _set_style(
    start: Range,
    end: Range,
    name: str,
    *,
    border: bool = True,
    gray: bool = False,
    font: bool = True,
    fill: bool = True,
    font_size: int | None = None,
) -> None:
    rng = start.sheet.range(start, end)

    if border:
        set_border(rng, edge_color=rcParams["frame.gray.border.color"] if gray else 0)

    if fill:
        _set_style_fill(rng, name, gray=gray)

    if font:
        _set_style_font(rng, name, gray=gray, font_size=font_size)


def _set_style_fill(rng: Range, name: str, *, gray: bool = False) -> None:
    if gray and name != "values":
        color = rcParams["frame.gray.fill.color"]
    else:
        color = rcParams[f"frame.{name}.fill.color"]

    set_fill(rng, color=color)


def _set_style_font(
    rng: Range,
    name: str,
    *,
    gray: bool = False,
    font_size: int | None = None,
) -> None:
    if gray:
        color = rcParams["frame.gray.font.color"]
    else:
        color = rcParams[f"frame.{name}.font.color"]
    bold = rcParams[f"frame.{name}.font.bold"]
    size = font_size or rcParams["frame.font.size"]

    set_font(rng, color=color, bold=bold, size=size)


@suspend_screen_updates
def set_frame_style(
    sf: SheetFrame,
    *,
    alignment: str | None = "center",
    banding: bool = False,
    succession: bool = False,
    border: bool = True,
    gray: bool = False,
    font: bool = True,
    fill: bool = True,
    font_size: int | None = None,
) -> None:
    """Set style of SheetFrame.

    Args:
        sf: The SheetFrame object.
        autofit: Whether to autofit the frame.
        alignment: The alignment of the frame.
        border: Whether to draw the border.
        font: Whether to specify the font.
        fill: Whether to fill the frame.
        banding: Whether to draw the banding.
        succession: Whether to hide the succession of the index.
        gray: Whether to set the frame in gray mode.
        font_size: The font size to specify directly.
    """
    cell = sf.cell
    sheet = sf.sheet

    set_style = partial(
        _set_style,
        border=border,
        gray=gray,
        font=font,
        fill=fill,
        font_size=font_size,
    )

    index_nlevels = sf.index.nlevels
    columns_nlevels = sf.columns.nlevels
    length = len(sf)

    start = cell
    end = cell.offset(columns_nlevels - 1, index_nlevels - 1)
    set_style(start, end, "index.name")

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

    if columns_nlevels > 1:
        start = cell.offset(0, index_nlevels)
        end = cell.offset(columns_nlevels - 2, index_nlevels + width - 1)
        set_style(start, end, "columns.name")

    start = cell.offset(columns_nlevels - 1, index_nlevels)
    end = cell.offset(columns_nlevels - 1, index_nlevels + width - 1)
    set_style(start, end, "columns")

    start = cell.offset(columns_nlevels, index_nlevels)
    end = cell.offset(columns_nlevels + length - 1, index_nlevels + width - 1)
    set_style(start, end, "values")

    rng = sheet.range(start, end)

    if banding and not gray:
        set_banding(rng)

    rng = sheet.range(cell, end)

    if border:
        if gray:
            ew = 2
            ec = rcParams["frame.gray.border.color"]
        else:
            ew = rcParams["frame.border.weight"]
            ec = rcParams["frame.border.color"]

        set_border(rng, edge_weight=ew, inside_weight=0, edge_color=ec)

    if alignment:
        set_alignment(rng, alignment)


def set_wide_column_style(sf: SheetFrame, gray: bool = False) -> None:
    edge_color = rcParams["frame.gray.border.color"] if gray else 0

    columns = list(sf.columns.wide_index)

    ew_default = rcParams["frame.wide-columns.border.weight"]

    for column in columns:
        loc = sf.columns.get_loc(column, sf.column + sf.index.nlevels)
        if not isinstance(loc, tuple):
            raise NotImplementedError

        rng = sf.sheet.range((sf.row, loc[0]), (sf.row, loc[1]))

        er = ew_default if column == columns[-1] else 2
        edge_weight = (1, er - 1, 1, 1) if gray else (2, er, 2, 2)
        set_border(rng, edge_weight, inside_weight=1, edge_color=edge_color)

        _set_style_fill(rng, "wide-columns", gray=gray)
        _set_style_font(rng, "wide-columns", gray=gray)

        rng = sf.sheet.range((sf.row - 1, loc[0]), (sf.row - 1, loc[1]))

        el = ew_default if column == columns[0] else 2
        edge_weight = (el - 1, 2, 2, 1) if gray else (el, ew_default, ew_default, 2)
        set_border(rng, edge_weight, inside_weight=0, edge_color=edge_color)

        _set_style_fill(rng, "wide-columns.name", gray=gray)
        _set_style_font(rng, "wide-columns.name", gray=gray)


def set_table_style(
    table: Table,
    even_color: int | str = rgb(240, 250, 255),
    odd_color: int | str = rgb(255, 255, 255),
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
    alignment: str | None = "center",
    border: bool = True,
    font: bool = True,
    fill: bool = True,
    font_size: int | None = None,
) -> None:
    """Set style of SheetFrame.

    Args:
        sf: The SheetFrame object.
        alignment: The alignment of the frame.
        border: Whether to draw the border.
        font: Whether to specify the font.
        fill: Whether to fill the frame.
        font_size: The font size to specify directly.
    """
    cell = sf.cell
    sheet = sf.sheet

    set_style = partial(
        _set_style,
        border=border,
        font=font,
        fill=fill,
        gray=False,
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
    _set_border(sf)


def _merge_index(index: Index, row: int, column: int, axis: int, sheet: Sheet) -> None:
    for start, end in iter_group_locs(index):
        if start == end:
            continue
        if axis == 0:
            sheet.range((row + start + 1, column), (row + end + 1, column)).merge()
        else:
            sheet.range((row, column + start + 1), (row, column + end + 1)).merge()


def _set_border(sf: HeatFrame) -> None:
    r = sf.row + 1
    c = sf.column + 1

    ec = rcParams["heat.border.color"]

    for row in iter_group_locs(sf.index, offset=r):
        if row[0] == row[1]:
            continue

        for col in iter_group_locs(sf.columns, offset=c):
            if col[0] == col[1]:
                continue

            rng = sf.sheet.range((row[0], col[0]), (row[1], col[1]))
            set_border(rng, edge_weight=2, edge_color=ec, inside_weight=0)
