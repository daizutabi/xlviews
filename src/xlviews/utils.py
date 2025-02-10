from __future__ import annotations

from typing import TYPE_CHECKING

import xlwings as xw
from pandas import Series
from xlwings import Range
from xlwings.constants import DVType, FormatConditionOperator

from xlviews.colors import rgb

if TYPE_CHECKING:
    from collections.abc import Iterator

    from numpy.typing import NDArray
    from pandas import DataFrame, Index
    from xlwings._xlwindows import COMRetryObjectWrapper

    from xlviews.dataframes.sheet_frame import SheetFrame


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


def iter_columns(sf: DataFrame | SheetFrame, columns: str | list[str]) -> Iterator[str]:
    """Yield the columns in the order of appearance with colon notation.

    Examples:
        >>> from pandas import DataFrame
        >>> df = DataFrame([[1, 2, 3]], columns=["A", "B", "C"])
        >>> list(iter_columns(df, "B"))
        ['B']
        >>> list(iter_columns(df, ":B"))
        ['A', 'B']
        >>> list(iter_columns(df, "::B"))
        ['A']
    """
    if isinstance(columns, str):
        columns = [columns]

    cs = [c for c in sf if isinstance(c, str)]

    for c in columns:
        if c.startswith("::"):
            yield from cs[: cs.index(c[2:])]
        elif c.startswith(":"):
            yield from cs[: cs.index(c[1:]) + 1]
        else:
            yield c


def iter_group_ranges(
    index: list | NDArray | Index | Series,
    padding: int = 0,
) -> Iterator[tuple[int, int]]:
    """Yield the group ranges of the index.

    The padding is added to the start and end of each group.
    The end is inclusive.

    Args:
        index: The index to iterate over.
        padding: The padding to add to the start and end of each group.

    Examples:
        >>> list(iter_group_ranges([1, 1, 1, 2, 2, 3, 3, 3]))
        [(0, 2), (3, 4), (5, 7)]

        >>> list(iter_group_ranges([1, 1, 1, 2, 2, 3, 3, 3], padding=3))
        [(0, 2), (6, 7), (11, 13)]
    """
    s = Series(index)
    idx = s[~s.duplicated()].index

    it = zip(idx, [*idx[1:], len(s)], strict=True)
    for k, (start, end) in enumerate(it):
        yield (start + padding * k, end + padding * k - 1)


def set_font_api(
    api: COMRetryObjectWrapper,
    name: str | None = None,
    *,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: int | str | None = None,
) -> None:
    font = api.Font
    if name:
        font.Name = name  # type: ignore
    if size:
        font.Size = size  # type: ignore
    if bold is not None:
        font.Bold = bold  # type: ignore
    if italic is not None:
        font.Italic = italic  # type: ignore
    if color is not None:
        font.Color = rgb(color)  # type: ignore


def add_validate_list(
    rng: Range,
    value: list[object],
    default: object | None = None,
) -> None:
    if default:
        rng.value = default

    type_ = DVType.xlValidateList
    operator = FormatConditionOperator.xlEqual
    formula = ",".join(map(str, value))

    rng.api.Validation.Add(Type=type_, Operator=operator, Formula1=formula)


# def label_func_from_list(columns, post=None):
#     """
#     カラム名のリストからラベル関数を作成して返す。

#     Parameters
#     ----------
#     columns : list of str
#         カラム名のリスト
#     post : str, optional
#         追加文字列

#     Returns
#     -------
#     callable
#     """

#     def get_format(t):
#         name_ = f"column.label.{t}"
#         if name_ in rcParams:
#             return rcParams[name_]
#         return "{" + t + "}"

#     fmt_dict = OrderedDict()
#     for column in columns:
#         fmt_dict[column] = get_format(column)

#     def func(**by_key):
#         labels = []
#         for by, fmt in fmt_dict.items():
#             key = by_key[by]
#             if isinstance(fmt, str):
#                 label = fmt.format(**{by: key})
#             else:
#                 label = fmt(key)
#             labels.append(label)
#         return "_".join(labels) + ("_" + post if post else "")

#     return func


# def format_label(data, fmt, sel=None, default=None):
#     dict_ = default.copy() if default else {}
#     if callable(fmt):
#         for column in data.columns:
#             try:
#                 values = data[column]
#             except TypeError:
#                 continue
#             if sel is not None:
#                 values = values[sel]
#             values = values.unique()
#             if len(values) == 1:
#                 dict_[column] = values[0]
#         return fmt(**dict_)
#     keys = re.findall(r"{([\w.]+)(?:}|:)", fmt)
#     for column in keys:
#         if column in data.columns:
#             values = data[column]
#             if sel is not None:
#                 values = values[sel]
#             values = values.unique()
#             if len(values) == 1:
#                 dict_[column] = values[0]
#     for key in keys:
#         if key not in dict_:
#             warnings.warn(
#                 f"タイトル文字列に含まれる'{key}'が、dfに含まれないか、単一ではない。",
#             )
#             dict_[key] = "XXX"
#     return fmt.format(**dict_)


# def get_sheet(book, name):
#     try:
#         return book.sheets[name]
#     except Exception:
#         return book.sheets.add(name, after=book.sheets(book.sheets.count))

# def get_range(book, name, title=False):
#     for sheet in book.sheets:
#         try:
#             range_ = sheet.names(name).refers_to_range
#             if title:
#                 start = range_[0, 0].offset(-1, 0)
#                 if start.value:
#                     return sheet.range(start, range_[-1, -1])
#             else:
#                 return range_
#         except Exception:
#             continue


# def copy_range(book_from, sheet_to, name, title=False):
#     range_ = get_range(book_from, name.replace("-", "__"), title=title)
#     range_.api.CopyPicture()  # Appearance:=xlScreen, Format:=xlPicture)
#     # sheet_to.activate()
#     # sheet_to.range('A1').api.Select()
#     sheet_to.api.Paste()
#     sheet_to.pictures[-1].name = name.replace("__", "-")

# def get_chart(book, name):
#     for sheet in book.sheets:
#         try:
#             return sheet.charts(name)
#         except Exception:
#             continue


# def copy_chart(book_from, sheet_to, name):
#     chart = get_chart(book_from, name)
#     # chart.api[1].ChartArea.Copy()
#     chart.api[0].Copy()
#     # sheet_to.api.Paste()
#     # sheet_to.activate()
#     # sheet_to.range('A1').api.Select()
#     sheet_to.api.PasteSpecial(Format="図 (PNG)", Link=False, DisplayAsIcon=False)
#     sheet_to.pictures[-1].name = name

# def outline_group(sheet, start: int, end: int, axis=0):
#     """
#     セルをグループする。
#     """
#     outline = sheet.api.Outline
#     if axis == 0:
#         outline.SummaryRow = constant("SummaryRow.xlSummaryAbove")
#         sheet.range((start, 1), (end, 1)).api.EntireRow.Group()
#     else:
#         outline.SummaryColumn = constant("SummaryColumn.xlSummaryOnLeft")
#         sheet.range((1, start), (1, end)).api.EntireRow.Group()


# def show_group(start: int, axis=0, show=True):
#     app = xw.apps.active
#     if axis == 0:
#         app.api.ExecuteExcel4Macro(f"SHOW.DETAIL(1,{start},{show})")
#     else:
#         raise ValueError("未実装")


# def hide_group(start: int, axis=0):
#     show_group(start, axis=axis, show=False)


# def outline_levels(sheet, levels: int, axis=0):
#     if axis == 0:
#         sheet.api.Outline.ShowLevels(RowLevels=levels)
#     else:
#         sheet.api.Outline.ShowLevels(ColumnLevels=levels)
