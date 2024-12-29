from collections import OrderedDict
from typing import Dict

NONCONST_VALUE = 'XXX'


def const(range_):
    """
    start-endの中の値が一意であればそれを返し，そうでなければXXXを返す数式を返す．
    """
    column = range_.get_address(column_absolute=False)
    ref = range_[0].offset(-1).get_address(column_absolute=False)
    subtotal = f'SUBTOTAL(3,{column})'
    column_name = (f'SUBSTITUTE(ADDRESS(ROW({column}),COLUMN({column}),4),'
                   f'ROW({column}),"")')
    value = (f'INDEX({column},MATCH(1,INDEX(SUBTOTAL(3,'
             f'INDIRECT({column_name}&ROW({column}))),),0))')
    prod_first = (f'SUBTOTAL(3,OFFSET({ref},'
                  f'ROW(INDIRECT("1:"&ROWS({column}))),))')
    prod_second = f'({column}={value})'
    sumproduct = f'SUMPRODUCT({prod_first}*{prod_second})'
    return f'IF({subtotal}={sumproduct},{value},"{NONCONST_VALUE}")'


AGGREGATE_FUNCTION = OrderedDict([('median', 12), ('soa', 999), ('count', 2),
                                  ('min', 5), ('mean', 1), ('max', 4),
                                  ('std', 8), ('sum', 9)])
AGGREGATE_FUNCTION_: Dict[str, int] = OrderedDict()


def _sort():
    for key in sorted(AGGREGATE_FUNCTION.keys()):
        AGGREGATE_FUNCTION_[key] = AGGREGATE_FUNCTION[key]


_sort()


# option=7: 非表示の行とエラー値を無視
def aggregate(func, *ranges, option=7, **kwargs):
    def get_address(range_, include_sheetname=False, **kwargs_):
        if hasattr(range_, 'get_address'):
            return range_.get_address(
                include_sheetname=include_sheetname, **kwargs_)
        else:
            sheetname = range_.Parent.Name
            range_ = range_.Address
            if include_sheetname:
                range_ = ','.join(
                    [f'{sheetname}!{range_}' for range_ in range_.split(',')])
            return range_

    column = ','.join([get_address(range_, **kwargs) for range_ in ranges])
    if func == 'soa':
        median = aggregate('median', *ranges, option=option, **kwargs)
        std = aggregate('std', *ranges, option=option, **kwargs)
        return f'{std}/{median}'
    elif func in AGGREGATE_FUNCTION:
        func = AGGREGATE_FUNCTION[func]
        return f'AGGREGATE({func},{option},{column})'
    elif hasattr(func, 'get_address'):  # funcの参照先は'mean'等の文字列
        ref = func.get_address(column_absolute=False, row_absolute=False)
        funcs, keys = zip(*AGGREGATE_FUNCTION_.items())
        keys = [f'{key}' for key in keys]
        keys = ','.join(keys)
        funcs = [f'"{func}"' for func in funcs]
        funcs = ','.join(funcs)
        soa = aggregate('soa', *ranges, option=option, **kwargs)
        func = f'LOOKUP({ref},{{{funcs}}},{{{keys}}})'
        return f'IF({ref}="soa",{soa},AGGREGATE({func},{option},{column}))'


def match_index(ref,
                sf,
                columns,
                column=None,
                na=False,
                null=False,
                error=False):
    """
    複数条件にマッチするインデックス(列番号 or 行番号，絶対)を返す数式文字列．

    Parameters
    ----------
    ref : xlviews.SheetFrame
        検索対象シートフレーム
    sf : xlviews.SheetFrame
        検索値を持つシートフレーム
    columns : str or list of str
        検索列
    column : str
        検索したシートフレームからピックアップするカラムを指定する．
    na : bool
        Trueのとき，エラーをNA()で置き換える．
    null : bool
        Trueのとき，エラーを""で置き換える．
    error
        False以外のとき，errorでエラーを置き換える

    Returns
    -------
    formula : str
        数式文字列
    """
    if isinstance(columns, str):
        columns = [columns]

    def gen():
        sf_columns = sf.columns
        for column_ in columns:
            if column_ in sf_columns:
                yield sf.range(column_), False
            else:
                yield sf.range(column_, 0)[0], True

    values, is_wides = zip(*gen())
    ranges = [ref.range(column_, -1) for column_ in columns]

    include_sheetname = ranges[0].sheet != values[0].sheet
    address = 'COLUMN' if len(ranges[0].rows) == 1 else 'ROW'

    conditions = []
    for k, (range_, value, is_wide) in enumerate(
            zip(ranges, values, is_wides)):
        range_ = range_.get_address(include_sheetname=include_sheetname)
        if k == 0:
            conditions.append(f'{address}({range_})')
        if is_wide:
            value = value.get_address(column_absolute=False)
        else:
            value = value.get_address(row_absolute=False)
        condition = '='.join([range_, value])
        condition = f'({condition})'
        conditions.append(condition)
    condition = '*'.join(conditions)
    formula = f'SUMPRODUCT({condition})'

    if column:
        cell = ref.range(column, 0)
        cell = cell.get_address(include_sheetname=include_sheetname)
        formula = (f'INDIRECT(ADDRESS({formula},COLUMN({cell}),1,1,' +
                   f'"{ref.sheet.name}"))')
        if error is not False:
            formula = f'IFERROR({formula},{error})'
        elif na:
            formula = f'IFERROR({formula},NA())'
        elif null:
            formula = f'IFERROR({formula},"")'
    return formula


def interp1d(x, y, value, error='""'):
    """
    xの範囲とyの範囲を線形補完する．xは昇順になっていること．

    Parameters
    ----------
    x : xlwings.Range
        xの範囲
    y : str or int
        yのカラム指定
    value : xlwings.Range
        新しいx
    error : str
        エラー時の値

    Returns
    -------
    formula : str
        数式文字列
    """
    include_sheetname = x.sheet != value.sheet

    def get_address(range_):
        return range_.get_address(include_sheetname=include_sheetname)

    value = value.get_address(row_absolute=False)
    xstart = get_address(x[0])
    xend = get_address(x[-1])
    ystart = f'INDIRECT(ADDRESS(ROW({xstart}),{y},1,1,"{x.sheet.name}"))'
    x = get_address(x)
    pre = f'AND({xstart}<={value},{value}<={xend})'
    match = f'MATCH({value},{x})'
    x = f'OFFSET({xstart},{match}-1,,2)'
    y = f'OFFSET({ystart},{match}-1,,2)'
    return f'IF({pre},TREND({y},{x},{value}),{error})'


def linear_fit(sf, x, y, to=None, a='a', b='b', by=None):
    """
    線形フィッティングを求める．

    Parameters
    ----------
    sf : xlviews.SheetFrame
    x, y: str
        カラム名
    to : xlviews.SheetFrame
        結果を記入するSheetFrame
    a, b: str
        カラム名
    by : str or list of str
        グルーピング
    """
    grouped = sf.groupby(by)
    xindex = sf.index(x)
    yindex = sf.index(y)
    if to is not None:
        if a not in to.columns:
            to[a] = 0
        a = to.range(a)
        if b not in to.columns:
            to[b] = 0
        b = to.range(b)

    for k, value in enumerate(grouped.values()):
        if len(value) != 1:
            raise ValueError('連続範囲のみ可能')
        x = sf.sheet.range((value[0][0], xindex), (value[0][1], xindex))
        y = sf.sheet.range((value[0][0], yindex), (value[0][1], yindex))
        if len(x) > 1:
            x = x.get_address()
            y = y.get_address()
            formula = f'IFERROR(SLOPE({y},{x}),NA())'
            a.offset(k).value = '=' + formula
            formula = f'IFERROR(INTERCEPT({y},{x}),NA())'
            b.offset(k).value = '=' + formula
        else:
            y = y.get_address()
            a.offset(k).value = '0'
            b.offset(k).value = '=' + y


def main():
    import xlviews as xv

    sf = xv.SheetFrame(2, 2, style=False, index_level=1)
    to = xv.SheetFrame(2, 6, style=False, index_level=0)
    linear_fit(sf, 'x', 'y', to, by='k')
    # columns = ['time', 'soa%']
    # y = match_index(sf, ref, columns=columns)
    # x = ref.column_range(0, -1)
    # value = sf.column_range('rate')
    # formula = interp1d(x, y, value)
    # sf['delta'] = '=' + formula


if __name__ == '__main__':
    main()
