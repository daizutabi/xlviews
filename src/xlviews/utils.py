from __future__ import annotations

from functools import wraps
from typing import TYPE_CHECKING, ParamSpec, TypeVar

import xlwings
from pandas import DataFrame, Series
from xlwings.constants import DVType, FormatConditionOperator

if TYPE_CHECKING:
    from collections.abc import Callable, Iterable, Iterator

    from numpy.typing import NDArray
    from pandas import Index
    from xlwings import Range as RangeImpl

    from xlviews.core.range import Range


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

    type_ = getattr(xlwings.constants, type_)

    return getattr(type_, name)


def iter_columns(
    iterable: Iterable | DataFrame,
    columns: str | list[str],
) -> Iterator[str]:
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

    lst = [x for x in iterable if isinstance(x, str)]

    for c in columns:
        if c.startswith("::"):
            yield from lst[: lst.index(c[2:])]
        elif c.startswith(":"):
            yield from lst[: lst.index(c[1:]) + 1]
        else:
            yield c


def iter_group_locs(
    index: list | NDArray | Index | Series,
    offset: int = 0,
    padding: int = 0,
) -> Iterator[tuple[int, int]]:
    """Yield the group ranges of the index.

    The padding is added to the start and end of each group.
    The end is inclusive.

    Args:
        index: The index to iterate over.
        padding: The padding to add to the start and end of each group.

    Examples:
        >>> list(iter_group_locs([1, 1, 1, 2, 2, 3, 3, 3]))
        [(0, 2), (3, 4), (5, 7)]

        >>> list(iter_group_locs([1, 1, 1, 2, 2, 3, 3, 3], offset=1))
        [(1, 3), (4, 5), (6, 8)]

        >>> list(iter_group_locs([1, 1, 1, 2, 2, 3, 3, 3], padding=3))
        [(0, 2), (6, 7), (11, 13)]
    """
    s = Series(index)
    idx = s[~s.duplicated()].index

    it = zip(idx, [*idx[1:], len(s)], strict=True)
    for k, (start, end) in enumerate(it):
        yield (start + offset + padding * k, end + offset + padding * k - 1)


def add_validate_list(
    rng: Range | RangeImpl,
    value: list[object],
    default: object | None = None,
) -> None:
    if default:
        rng.value = default

    type_ = DVType.xlValidateList
    operator = FormatConditionOperator.xlEqual
    formula = ",".join(map(str, value))

    rng.api.Validation.Add(Type=type_, Operator=operator, Formula1=formula)


P = ParamSpec("P")
R = TypeVar("R")


def suspend_screen_updates(func: Callable[P, R]) -> Callable[P, R]:
    """Suspend screen updates to speed up operations."""

    @wraps(func)
    def _func(*args: P.args, **kwargs: P.kwargs) -> R:
        if app := xlwings.apps.active:
            is_updating = app.screen_updating
            app.screen_updating = False

        try:
            return func(*args, **kwargs)
        finally:
            if app:
                app.screen_updating = is_updating

    return _func


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
#                 f"タイトル文字列に含まれる'{key}'が、dfに含まれない。",
#             )
#             dict_[key] = "XXX"
#     return fmt.format(**dict_)
