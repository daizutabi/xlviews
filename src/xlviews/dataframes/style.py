"""Set styles for SheetFrame."""

from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING

import pywintypes
from xlwings.constants import TableStyleElementType

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.style import (
    hide_succession,
    hide_unique,
    set_alignment,
    set_banding,
    set_border,
    set_fill,
    set_font,
)
from xlviews.utils import rgb

if TYPE_CHECKING:
    from xlwings import Range

    from .sheet import SheetFrame
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


@turn_off_screen_updating
def set_frame_style(
    sf: SheetFrame,
    *,
    autofit: bool = False,
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

    index_level = sf.index_level
    columns_level = sf.columns_level
    length = len(sf)

    if index_level > 0:
        start = cell
        end = cell.offset(columns_level - 1, index_level - 1)
        set_style(start, end, "index.name")

        start = cell.offset(columns_level, 0)
        end = cell.offset(columns_level + length - 1, index_level - 1)
        set_style(start, end, "index")

        if succession:
            rng = sheet.range(start.offset(1, 0), end)
            hide_succession(rng)

            start = cell.offset(columns_level - 1, 0)
            end = cell.offset(columns_level - 1, index_level - 1)
            rng = sheet.range(start, end)
            hide_unique(rng, length)

    width = len(sf.value_columns)

    if columns_level > 1:
        start = cell.offset(0, index_level)
        end = cell.offset(columns_level - 2, index_level + width - 1)
        set_style(start, end, "columns.name")

    start = cell.offset(columns_level - 1, index_level)
    end = cell.offset(columns_level - 1, index_level + width - 1)
    set_style(start, end, "columns")

    start = cell.offset(columns_level, index_level)
    end = cell.offset(columns_level + length - 1, index_level + width - 1)
    set_style(start, end, "values")

    rng = sheet.range(start, end)

    if banding and not gray:
        set_banding(rng)

    rng = sheet.range(cell, end)

    if border:
        ew = 2 if gray else 3
        ec = rcParams["frame.gray.border.color"] if gray else 0
        set_border(rng, edge_weight=ew, inside_weight=0, edge_color=ec)

    if autofit:
        rng.columns.autofit()

    if alignment:
        set_alignment(rng, alignment)


def set_wide_column_style(sf: SheetFrame, gray: bool = False) -> None:
    wide_columns = sf.wide_columns
    edge_color = rcParams["frame.gray.border.color"] if gray else 0

    for wide_column in wide_columns:
        rng = sf.range(wide_column, 0).offset(-1)

        er = 3 if wide_column == wide_columns[-1] else 2
        edge_weight = (1, er - 1, 1, 1) if gray else (2, er, 2, 2)
        set_border(rng, edge_weight, inside_weight=1, edge_color=edge_color)

        _set_style_fill(rng, "wide-columns", gray=gray)
        _set_style_font(rng, "wide-columns", gray=gray)

    for wide_column in wide_columns:
        rng = sf.range(wide_column, 0).offset(-2)

        el = 3 if wide_column == wide_columns[0] else 2
        edge_weight = (el - 1, 2, 2, 1) if gray else (el, 3, 3, 2)
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
