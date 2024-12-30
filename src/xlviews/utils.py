from __future__ import annotations

import re
import warnings
from collections import OrderedDict
from collections.abc import Hashable
from typing import TYPE_CHECKING

import matplotlib as mpl
import xlwings as xw
from pandas import DataFrame, Series
from xlwings import Range, Sheet

from xlviews.config import rcParams

if TYPE_CHECKING:
    import numpy as np
    from numpy.typing import NDArray


def constant(type_: str, name: str | None = None) -> int:
    """Return the Excel constant.

    Args:
        type_ (str): The type name.
        name (str): The name.

    Examples:
        >>> constant("BordersIndex", "EdgeTop")
        8
    """
    if name is None:
        if "." in type_:
            type_, name = type_.split(".")
        else:
            type_, name = "Constants", type_

    if not name.startswith("xl"):
        name = "xl" + name[0].upper() + name[1:]

    type_ = getattr(xw.constants, type_)

    return getattr(type_, name)


def rgb(
    color: int | tuple[int, int, int] | str,
    green: int | None = None,
    blue: int | None = None,
) -> int:
    """Return the RGB color integer.

    Args:
        color (int, tuple[int, int, int], or str): The color or red value.
        green (int): The green value.
        blue (int): The blue value.

    Examples:
        >>> rgb(4)
        4

        >>> rgb((100, 200, 40))
        2672740

        >>> rgb("pink")
        13353215

        >>> rgb("#123456")
        5649426
    """
    if isinstance(color, int) and green is None and blue is None:
        return color

    if all(isinstance(x, int) for x in [color, green, blue]):
        return color + green * 256 + blue * 256 * 256  # type: ignore

    if isinstance(color, str):
        color = mpl.colors.cnames.get(color, color)  # type: ignore

        if not isinstance(color, str) or not color.startswith("#") or len(color) != 7:
            raise ValueError("Invalid color format. Expected #xxxxxx.")

        return rgb(int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16))

    if isinstance(color, tuple):
        return rgb(*color)

    raise ValueError("Invalid color format. Expected #xxxxxx.")


def array_index(
    values: list | NDArray | DataFrame | Series,
    sel: list[bool] | NDArray[np.bool_] | None = None,
) -> dict[Hashable, list[list[int]]]:
    """Return a dictionary indicating the positions where values exist.

    The keys of the dictionary are the values in `values`. The values of the
    dictionary are lists of positions where the keys exist, in the format:

        [[start1, end1], [start2, end2], ...]

    The `end` is the inclusive position of the element, different from slice
    notation.

    Args:
        values (listable): Array to scan for value positions.
        sel (list of bool, optional): Specifies whether to detect values or not.
            Indices where this value is False are excluded.

    Returns:
        dict: Dictionary storing the positions where values exist.

    Examples:
        >>> values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
        >>> index = array_index(values)
        >>> index[(1, 2)]
        [[0, 1], [4, 4]]
        >>> index[(3, 4)]
        [[2, 3], [5, 6]]

        >>> sel = [True, False, True, False, True, False, False]
        >>> index = array_index(values, sel=sel)
        >>> index[(1, 2)]
        [[0, 0], [4, 4]]
        >>> index[(3, 4)]
        [[2, 2]]
    """
    if len(values) == 0:
        return {}

    if isinstance(values, DataFrame | Series):
        values = values.to_numpy()

    if not isinstance(values[0], Hashable):
        values = [tuple(x) for x in values]

    index: dict[Hashable, list[list[int]]] = {}
    for k, x in enumerate(values):
        if sel is not None and not sel[k]:
            continue

        if x not in index:
            index[x] = [[k, k]]
        else:
            current = index[x]
            if k == current[-1][-1] + 1:
                current[-1][-1] = k
            else:
                current.append([k, k])

    return index


def get_sheet(book, name):
    try:
        return book.sheets[name]
    except Exception:
        return book.sheets.add(name, after=book.sheets(book.sheets.count))


def get_chart(book, name):
    for sheet in book.sheets:
        try:
            return sheet.charts(name)
        except Exception:
            continue


def get_range(book, name, title=False):
    for sheet in book.sheets:
        try:
            range_ = sheet.names(name).refers_to_range
            if title:
                start = range_[0, 0].offset(-1, 0)
                if start.value:
                    return sheet.range(start, range_[-1, -1])
            else:
                return range_
        except Exception:
            continue


def copy_chart(book_from, sheet_to, name):
    chart = get_chart(book_from, name)
    # chart.api[1].ChartArea.Copy()
    chart.api[0].Copy()
    # sheet_to.api.Paste()
    # sheet_to.activate()
    # sheet_to.range('A1').api.Select()
    sheet_to.api.PasteSpecial(Format="図 (PNG)", Link=False, DisplayAsIcon=False)
    sheet_to.pictures[-1].name = name


def copy_range(book_from, sheet_to, name, title=False):
    range_ = get_range(book_from, name.replace("-", "__"), title=title)
    range_.api.CopyPicture()  # Appearance:=xlScreen, Format:=xlPicture)
    # sheet_to.activate()
    # sheet_to.range('A1').api.Select()
    sheet_to.api.Paste()
    sheet_to.pictures[-1].name = name.replace("__", "-")


def add_validation(cell, value, default=None):
    if default:
        cell.value = default
    if isinstance(value, list):
        type_ = constant("DVType.xlValidateList")
        operator = constant("FormatConditionOperator.xlEqual")
        value = ",".join([str(x) for x in value])
    else:
        raise ValueError("未実装")

    cell.api.Validation.Add(Type=type_, Operator=operator, Formula1=value)


def outline_group(sheet, start: int, end: int, axis=0):
    """
    セルをグループする。
    """
    outline = sheet.api.Outline
    if axis == 0:
        outline.SummaryRow = constant("SummaryRow.xlSummaryAbove")
        sheet.range((start, 1), (end, 1)).api.EntireRow.Group()
    else:
        outline.SummaryColumn = constant("SummaryColumn.xlSummaryOnLeft")
        sheet.range((1, start), (1, end)).api.EntireRow.Group()


def show_group(start: int, axis=0, show=True):
    app = xw.apps.active
    if axis == 0:
        app.api.ExecuteExcel4Macro(f"SHOW.DETAIL(1,{start},{show})")
    else:
        raise ValueError("未実装")


def hide_group(start: int, axis=0):
    show_group(start, axis=axis, show=False)


def outline_levels(sheet, levels: int, axis=0):
    if axis == 0:
        sheet.api.Outline.ShowLevels(RowLevels=levels)
    else:
        sheet.api.Outline.ShowLevels(ColumnLevels=levels)


def label_func_from_list(columns, post=None):
    """
    カラム名のリストからラベル関数を作成して返す。

    Parameters
    ----------
    columns : list of str
        カラム名のリスト
    post : str, optional
        追加文字列

    Returns
    -------
    callable
    """

    def get_format(t):
        name_ = f"column.label.{t}"
        if name_ in rcParams:
            return rcParams[name_]
        return "{" + t + "}"

    fmt_dict = OrderedDict()
    for column in columns:
        fmt_dict[column] = get_format(column)

    def func(**by_key):
        labels = []
        for by, fmt in fmt_dict.items():
            key = by_key[by]
            if isinstance(fmt, str):
                label = fmt.format(**{by: key})
            else:
                label = fmt(key)
            labels.append(label)
        return "_".join(labels) + ("_" + post if post else "")

    return func


def format_label(data, fmt, sel=None, default=None):
    dict_ = default.copy() if default else {}
    if callable(fmt):
        for column in data.columns:
            try:
                values = data[column]
            except TypeError:
                continue
            if sel is not None:
                values = values[sel]
            values = values.unique()
            if len(values) == 1:
                dict_[column] = values[0]
        return fmt(**dict_)
    keys = re.findall(r"{([\w.]+)(?:}|:)", fmt)
    for column in keys:
        if column in data.columns:
            values = data[column]
            if sel is not None:
                values = values[sel]
            values = values.unique()
            if len(values) == 1:
                dict_[column] = values[0]
    for key in keys:
        if key not in dict_:
            warnings.warn(
                f"タイトル文字列に含まれる'{key}'が、"
                "dfに含まれないか、単一ではない。",
            )
            dict_[key] = "XXX"
    return fmt.format(**dict_)


def columns_list(df, columns):
    """
    ':column' or '::column' 形式を通常のカラム名のリストに変換する。
    ':column'はcolumnを含める、'::column'はcolumnの一つ前まで.
    文字列の場合はリストにする。

    Parameters
    ----------
    df : DataFrame or SheetFrame
    columns : str or list of str
        カラム名

    Returns
    -------
    columns : list of str
        カラム名のリスト

    """
    if isinstance(columns, str):
        columns = [columns]
    columns_ = list(df.columns)

    def gen():
        for column in columns:
            if column.startswith("::"):
                yield from columns_[: columns_.index(column[2:])]
            elif column.startswith(":"):
                yield from columns_[: columns_.index(column[1:]) + 1]
            else:
                yield column

    return list(gen())


def delete_charts(sheet=None):
    """
    シート内のすべてのチャートを削除する。

    Parameters
    ----------
    sheet : xlwings.Sheet
    """
    if sheet is None:
        sheet = xw.sheets.active
    for chart in sheet.charts:
        chart.delete()


def set_axis_dimension(key, label=None, ticks=None, format=None):
    if label:
        rcParams[f"axis.label.{key}"] = label
    if ticks:
        rcParams[f"axis.ticks.{key}"] = ticks
    if format:
        rcParams[f"axis.format.{key}"] = format


def autofilter(list_object, *args, **field_criteria):
    """
    キーワード引数で指定される条件に応じてフィルタリングする
    キーワード引数のキーはカラム名、値は条件。条件は以下のものが指定できる。
       - list : 要素を指定する。
       - tuple : 値の範囲を指定する。
       - None : 設定されているフィルタをクリアする
       - 他 : 値の一致

    """
    for field, criteria in zip(args[::2], args[1::2], strict=False):
        field_criteria[field] = criteria

    filter_ = list_object.Range.AutoFilter
    operator = xw.constants.AutoFilterOperator
    columns = [column.Name for column in list_object.ListColumns]

    for field, criteria in field_criteria.items():
        field_index = columns.index(field) + 1
        if isinstance(criteria, list):
            criteria = list(map(str, criteria))
            filter_(
                Field=field_index,
                Criteria1=criteria,
                Operator=operator.xlFilterValues,
            )
        elif isinstance(criteria, tuple):
            filter_(
                Field=field_index,
                Criteria1=f">={criteria[0]}",
                Operator=operator.xlAnd,
                Criteria2=f"<={criteria[1]}",
            )
        elif criteria is None:
            filter_(Field=field_index)
        else:
            filter_(Field=field_index, Criteria1=f"{criteria}")


def main():
    sheet = xw.sheets.active
    list_object = sheet.api.ListObjects("テーブル1")
    autofilter(list_object, TMR=(100, 150))
