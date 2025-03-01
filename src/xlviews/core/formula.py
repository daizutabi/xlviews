from __future__ import annotations

from typing import TYPE_CHECKING, TypeAlias

from xlwings import Range as RangeImpl

from .range import Range
from .range_collection import RangeCollection

if TYPE_CHECKING:
    from collections.abc import Iterable

Func: TypeAlias = str | Range | RangeImpl | None

NONCONST_VALUE = "*"


def const(rng: Range | RangeImpl) -> str:
    """Return a formula to check if the values in the range are unique."""
    column = rng.get_address(column_absolute=False)
    ref = rng[0].offset(-1).get_address(column_absolute=False)

    subtotal = f"SUBTOTAL(3,{column})"
    name = f'SUBSTITUTE(ADDRESS(ROW({column}),COLUMN({column}),4),ROW({column}),"")'
    index = f"INDEX(SUBTOTAL(3,INDIRECT({name}&ROW({column}))),)"
    value = f"INDEX({column},MATCH(1,{index},0))"

    prod_first = f'SUBTOTAL(3,OFFSET({ref},ROW(INDIRECT("1:"&ROWS({column}))),))'
    prod_second = f"({column}={value})"
    sumproduct = f"SUMPRODUCT({prod_first}*{prod_second})"

    return f'IFNA(IF({subtotal}={sumproduct},{value},"{NONCONST_VALUE}"),"")'


AGG_FUNCS = {
    "median": 12,
    "soa": 999,
    "count": 2,
    "min": 5,
    "mean": 1,
    "max": 4,
    "std": 8,
    "sum": 9,
}

AGG_FUNCS_SORTED = {key: AGG_FUNCS[key] for key in sorted(AGG_FUNCS.keys())}
AGG_FUNC_NAMES = ",".join(f'"{name}"' for name in AGG_FUNCS_SORTED)
AGG_FUNC_INTS = ",".join(f'"{value}"' for value in AGG_FUNCS_SORTED.values())


def _aggregate(
    func: Func,
    ranges: Range | RangeCollection | Iterable[Range | RangeCollection] | str,
    option: int,
    **kwargs,
) -> str:
    if func == "soa":
        std = aggregate("std", ranges, option, **kwargs)
        median = aggregate("median", ranges, option, **kwargs)
        return f"{std}/{median}"

    if isinstance(ranges, Range | RangeCollection):
        ranges = [ranges]

    if isinstance(ranges, str):
        column = ranges
    else:
        column = ",".join(r.get_address(**kwargs) for r in ranges)

    if func is None:
        return column

    if isinstance(func, str):
        if func in AGG_FUNCS:
            return f"AGGREGATE({AGG_FUNCS[func]},{option},{column})"

        msg = f"Invalid aggregate function: {func}"
        raise ValueError(msg)

    ref = func.get_address(column_absolute=False, row_absolute=False)
    func = f"LOOKUP({ref},{{{AGG_FUNC_NAMES}}},{{{AGG_FUNC_INTS}}})"
    soa = aggregate("soa", ranges, option=option, **kwargs)
    return f'IF({ref}="soa",{soa},AGGREGATE({func},{option},{column}))'


def aggregate(
    func: Func,
    ranges: Range | RangeCollection | Iterable[Range | RangeCollection] | str,
    option: int = 7,  # ignore hidden rows and error values
    row_absolute: bool = True,
    column_absolute: bool = True,
    include_sheetname: bool = False,
    external: bool = False,
    formula: bool = False,
) -> str:
    value = _aggregate(
        func,
        ranges,
        option,
        row_absolute=row_absolute,
        column_absolute=column_absolute,
        include_sheetname=include_sheetname,
        external=external,
    )

    if formula:
        return f"={value}"

    return value


# def match_index(ref, sf, columns, column=None, na=False, null=False, error=False):
#     """
#     複数条件にマッチするインデックス(列番号 or 行番号、絶対)を返す数式文字列。

#     Parameters
#     ----------
#     ref : xlviews.SheetFrame
#         検索対象シートフレーム
#     sf : xlviews.SheetFrame
#         検索値を持つシートフレーム
#     columns : str or list of str
#         検索列
#     column : str
#         検索したシートフレームからピックアップするカラムを指定する。
#     na : bool
#         Trueのとき、エラーをNA()で置き換える。
#     null : bool
#         Trueのとき、エラーを""で置き換える。
#     error
#         False以外のとき、errorでエラーを置き換える

#     Returns
#     -------
#     formula : str
#         数式文字列
#     """
#     if isinstance(columns, str):
#         columns = [columns]

#     def gen():
#         sf_columns = sf.columns
#         for column_ in columns:
#             if column_ in sf_columns:
#                 yield sf.range(column_), False
#             else:
#                 yield sf.range(column_, 0)[0], True

#     values, is_wides = zip(*gen(), strict=False)
#     ranges = [ref.range(column_, -1) for column_ in columns]

#     include_sheetname = ranges[0].sheet != values[0].sheet
#     address = "COLUMN" if len(ranges[0].rows) == 1 else "ROW"

#     conditions = []
#     for k, (range_, value, is_wide) in enumerate(
#         zip(ranges, values, is_wides, strict=False),
#     ):
#         range_ = range_.get_address(include_sheetname=include_sheetname)
#         if k == 0:
#             conditions.append(f"{address}({range_})")
#         if is_wide:
#             value = value.get_address(column_absolute=False)
#         else:
#             value = value.get_address(row_absolute=False)
#         condition = "=".join([range_, value])
#         condition = f"({condition})"
#         conditions.append(condition)
#     condition = "*".join(conditions)
#     formula = f"SUMPRODUCT({condition})"

#     if column:
#         cell = ref.range(column, 0)
#         cell = cell.get_address(include_sheetname=include_sheetname)
#         formula = (
#             f"INDIRECT(ADDRESS({formula},COLUMN({cell}),1,1," + f'"{ref.sheet.name}"))'
#         )
#         if error is not False:
#             formula = f"IFERROR({formula},{error})"
#         elif na:
#             formula = f"IFERROR({formula},NA())"
#         elif null:
#             formula = f'IFERROR({formula},"")'
#     return formula


# def interp1d(x, y, value, error='""'):
#     """
#     xの範囲とyの範囲を線形補完する。xは昇順になっていること。

#     Parameters
#     ----------
#     x : xlwings.Range
#         xの範囲
#     y : str or int
#         yのカラム指定
#     value : xlwings.Range
#         新しいx
#     error : str
#         エラー時の値

#     Returns
#     -------
#     formula : str
#         数式文字列
#     """
#     include_sheetname = x.sheet != value.sheet

#     def get_address(range_):
#         return range_.get_address(include_sheetname=include_sheetname)

#     value = value.get_address(row_absolute=False)
#     xstart = get_address(x[0])
#     xend = get_address(x[-1])
#     ystart = f'INDIRECT(ADDRESS(ROW({xstart}),{y},1,1,"{x.sheet.name}"))'
#     x = get_address(x)
#     pre = f"AND({xstart}<={value},{value}<={xend})"
#     match = f"MATCH({value},{x})"
#     x = f"OFFSET({xstart},{match}-1,,2)"
#     y = f"OFFSET({ystart},{match}-1,,2)"
#     return f"IF({pre},TREND({y},{x},{value}),{error})"


# def linear_fit(sf, x, y, to=None, a="a", b="b", by=None):
#     """
#     線形フィッティングを求める。

#     Parameters
#     ----------
#     sf : xlviews.SheetFrame
#     x, y: str
#         カラム名
#     to : xlviews.SheetFrame
#         結果を記入するSheetFrame
#     a, b: str
#         カラム名
#     by : str or list of str
#         グルーピング
#     """
#     grouped = sf.groupby(by)
#     xindex = sf.index(x)
#     yindex = sf.index(y)
#     if to is not None:
#         if a not in to.columns:
#             to[a] = 0
#         a = to.range(a)
#         if b not in to.columns:
#             to[b] = 0
#         b = to.range(b)

#     for k, value in enumerate(grouped.values()):
#         if len(value) != 1:
#             raise ValueError("連続範囲のみ可能")
#         x = sf.sheet.range((value[0][0], xindex), (value[0][1], xindex))
#         y = sf.sheet.range((value[0][0], yindex), (value[0][1], yindex))
#         if len(x) > 1:
#             x = x.get_address()
#             y = y.get_address()
#             formula = f"IFERROR(SLOPE({y},{x}),NA())"
#             a.offset(k).value = "=" + formula
#             formula = f"IFERROR(INTERCEPT({y},{x}),NA())"
#             b.offset(k).value = "=" + formula
#         else:
#             y = y.get_address()
#             a.offset(k).value = "0"
#             b.offset(k).value = "=" + y


# def main():
#     import xlviews as xv

#     sf = xv.SheetFrame(2, 2, style=False, index.nlevels=1)
#     to = xv.SheetFrame(2, 6, style=False, index.nlevels=0)
#     linear_fit(sf, "x", "y", to, by="k")
#     # columns = ['time', 'soa%']
#     # y = match_index(sf, ref, columns=columns)
#     # x = ref.column_range(0, -1)
#     # value = sf.column_range('rate')
#     # formula = interp1d(x, y, value)
#     # sf['delta'] = '=' + formula


# if __name__ == "__main__":
#     main()
