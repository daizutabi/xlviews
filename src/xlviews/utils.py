import os
import re
import warnings
from collections import OrderedDict

import matplotlib
import pandas as pd
import xlwings as xw

from xlviews.config import rcParams


def constant(type_, name=None):
    """
    エクセル定数を返す．

    Parameters
    ----------
    type_ : str
        型名
    name : str
        名前

    Examples
    --------
    >>> constant('BordersIndex', 'EdgeTop')
    8
    """
    if name is None:
        if '.' in type_:
            type_, name = type_.split('.')
        else:
            type_, name = 'Constants', type_
    if not name.startswith('xl'):
        name = 'xl' + name[0].upper() + name[1:]
    type_ = getattr(xw.constants, type_)
    return getattr(type_, name)


def rgb(color, green=None, blue=None):
    """
    RGBの色関数．

    Parameters
    ----------
    color : int, list, or str
    green: int
    blue: int

    Examples
    --------
    >>> rgb(4)
    4
    >>> rgb([100, 200, 40])
    2672740
    >>> rgb('pink')
    13353215
    >>> rgb('#123456')
    5649426
    """
    if green is not None:
        red = color
    elif isinstance(color, int):
        return color
    elif isinstance(color, str):
        color = matplotlib.colors.cnames.get(color, color)
        if not color.startswith('#') or len(color) != 7:
            raise ValueError('colorが文字列のときは，#xxxxxxの形式．')
        red = int(color[1:3], 16)
        green = int(color[3:5], 16)
        blue = int(color[5:7], 16)
    else:
        red, green, blue = color

    return red + green * 256 + blue * 256 * 256


def array_index(values, sel=None):
    """
    値が存在する位置を辞書で返す．
    辞書のキーは，valuesの値．辞書の値は，そのキーが存在する位置のリストで，
    [[start1, end1], [start2, end2], ...]
    の形式．endは要素の位置そのもので，スライス表記とは異なる．

    Parameters
    ----------
    values : listable
        値の位置を走査する配列
    sel : list of bool, optional
        そもそも値を検出するかを指定する
        この値がFalseのindexは除外される．

    Returns
    -------
    dict
        値が存在する位置を格納した辞書．

    Examples
    --------
    >>> values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
    >>> dict_ = array_index(values)
    >>> dict_[(1, 2)]
    [[0, 1], [4, 4]]
    >>> dict_[(3, 4)]
    [[2, 3], [5, 6]]
    """
    dict_ = OrderedDict()
    if len(values) == 0:
        return dict_

    if isinstance(values, pd.DataFrame) or isinstance(values, pd.Series):
        values = values.values

    try:
        hash(values[0])
    except TypeError:
        values = [tuple(x) for x in values]

    for k, x in enumerate(values):
        if sel is not None and not sel[k]:
            continue
        if x not in dict_:
            dict_[x] = [[k, k]]
        else:
            index = dict_[x]
            if k == index[-1][-1] + 1:
                index[-1][-1] = k
            else:
                index.append([k, k])
    return dict_


def multirange(sheet, row, column):
    """
    飛び地Rangeを作成する．
    rowとcolumnのいずれかはint.
    intでない方をindexとしたとき，indexはlist.
    indexが(int, int)の場合は単純なRange
    それ以外のときは，indexの各要素は，int or (int, int)で
    それらを連結して，非連続なRangeを作製して返す．

    Parameters
    ----------
    sheet : xlwings.main.Sheet
        シートオブジェクト
    row : int or list
        行番号
    column : int or list
        列番号
    """
    if isinstance(row, int):
        axis = 0
        index = column  # type: list
    elif isinstance(column, int):
        axis = 1
        index = row  # type: list
    else:
        raise ValueError('rowとcolumnのどちらかはintでなければならない．')

    if isinstance(index, int):
        return sheet.range(row, column).api

    def _range(start_end):
        if isinstance(start_end, int):
            start = end = start_end
        else:
            start, end = start_end
        if axis == 0:
            return sheet.range((row, start), (row, end))
        else:
            return sheet.range((start, column), (end, column))

    if (len(index) == 2 and isinstance(index[0], int) and
            isinstance(index[1], int)):
        index = [index]

    ranges = [_range(i).api for i in index]
    union = sheet.book.app.api.Union
    range_ = ranges[0]
    for r in ranges[1:]:
        range_ = union(range_, r)

    return range_


def multirange_indirect(sheet, row, column):
    """
    不連続範囲でもSLOPE関数などが扱えるようにする．
    戻り値はstr
    """
    ranges = multirange(sheet, row, column)
    address = ','.join(['"' + range_.Address + '"' for range_ in ranges])
    return 'N(INDIRECT({' + address + '}))'


def reference(sheet, cell):
    """
    Sheetのセルへの参照を返す．
    cellが文字列であればそのまま返す．
    """
    if isinstance(cell, tuple):
        # TODO: tupleのときどの要素使う？連結する？
        cell = cell[0]
    if not isinstance(cell, str):
        cell = sheet.range(*cell).get_address(include_sheetname=True)
        cell = '=' + cell
    return cell


def open_or_create(path, app=None, sheetname=None, visible=True):
    """
    pathのエクセルファイルが存在すれば開く．なければ新規ワークブックを作成
    する．

    Parameters
    ----------
    path : str
        ファイルパス
    app : xw.App
        アプリケーションを指定する．
    sheetname : str
        シート名. 指定されるとSheetオブジェクトが返される．
    visible : bool
        可視かどうか

    Returns
    -------
    book_or_sheet : xw.Book, xw.Sheet
    """
    app = app or xw.apps.active
    if app is None:
        created = True
        app = xw.apps.add()
    else:
        created = False

    if os.path.exists(path):
        book = app.books.open(path)
        if created:
            app.books[0].close()
        created = False
    else:
        if created:
            book = app.books[0]
        else:
            book = app.books.add()
        book.save(path)
        created = True
    app.visible = visible

    if sheetname is None:
        return book

    if created:
        sheet = book.sheets[0]
        sheet.name = sheetname
        book.save()
        return sheet
    else:
        sheet = None
        for sheet in book.sheets:
            if sheet.name == sheetname:
                return sheet
        sheet = book.sheets.add(sheetname, after=sheet)
        book.save()
        return sheet


def get_sheet_cell_row_column(*args):
    """
    諸関数の位置引数に指定される引数からシート，セル，ロー，
    カラムを取得する．

    *argsに指定できる方法は以下の通り
      - sheet, row, column
      - sheet, (row, columm)
      - row, column
      - (row, column)
      - cell
      - ()
      - 'A1'
    """
    if not xw.apps:
        xw.apps.add()
        sheet = xw.sheets.active
        from xlviews.style import hide_gridlines
        hide_gridlines(sheet)
        sheet.range('A1').column_width = 1
        sheet.range('B2').select()
        if args and isinstance(args[0], str):
            sheet.name = args[0]
    else:
        sheet = xw.sheets.active

    if len(args) == 3:
        sheet, row, column = args
    elif len(args) == 2:
        if isinstance(args[0], int):
            row, column = args
        else:
            sheet, (row, column) = args
    elif len(args) == 1:
        if isinstance(args[0], str):
            cell = sheet.range(args[0])
            row, column = cell.row, cell.column
        elif isinstance(args[0], tuple):
            row, column = args[0]
        else:
            cell = args[0]
            sheet = cell.sheet
            row, column = cell.row, cell.column
    elif len(args) == 0:
        cell = sheet.book.selection
        row, column = cell.row, cell.column
    else:
        raise ValueError('引数の長さが4以上', len(args))

    if isinstance(sheet, str):
        book = xw.books.active
        try:
            sheet = book.sheets(sheet)
        except Exception:
            sheet = book.sheets.add(sheet, after=book.sheets[-1])

    cell = sheet.range(row, column)
    return sheet, cell, row, column


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
    sheet_to.api.PasteSpecial(Format='図 (PNG)', Link=False,
                              DisplayAsIcon=False)
    sheet_to.pictures[-1].name = name


def copy_range(book_from, sheet_to, name, title=False):
    range_ = get_range(book_from, name.replace('-', '__'), title=title)
    range_.api.CopyPicture()  # Appearance:=xlScreen, Format:=xlPicture)
    # sheet_to.activate()
    # sheet_to.range('A1').api.Select()
    sheet_to.api.Paste()
    sheet_to.pictures[-1].name = name.replace('__', '-')


def add_validation(cell, value, default=None):
    if default:
        cell.value = default
    if isinstance(value, list):
        type_ = constant('DVType.xlValidateList')
        operator = constant('FormatConditionOperator.xlEqual')
        value = ','.join([str(x) for x in value])
    else:
        raise ValueError('未実装')

    cell.api.Validation.Add(Type=type_, Operator=operator, Formula1=value)


def outline_group(sheet, start: int, end: int, axis=0):
    """
    セルをグループする．
    """
    outline = sheet.api.Outline
    if axis == 0:
        outline.SummaryRow = constant('SummaryRow.xlSummaryAbove')
        sheet.range((start, 1), (end, 1)).api.EntireRow.Group()
    else:
        outline.SummaryColumn = constant('SummaryColumn.xlSummaryOnLeft')
        sheet.range((1, start), (1, end)).api.EntireRow.Group()


def show_group(start: int, axis=0, show=True):
    app = xw.apps.active
    if axis == 0:
        app.api.ExecuteExcel4Macro(f'SHOW.DETAIL(1,{start},{show})')
    else:
        raise ValueError('未実装')


def hide_group(start: int, axis=0):
    show_group(start, axis=axis, show=False)


def outline_levels(sheet, levels: int, axis=0):
    if axis == 0:
        sheet.api.Outline.ShowLevels(RowLevels=levels)
    else:
        sheet.api.Outline.ShowLevels(ColumnLevels=levels)


def label_func_from_list(columns, post=None):
    """
    カラム名のリストからラベル関数を作成して返す．

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
        name_ = f'column.label.{t}'
        if name_ in rcParams:
            return rcParams[name_]
        else:
            return '{' + t + '}'

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
        return '_'.join(labels) + ('_' + post if post else '')

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
    else:
        keys = re.findall(r'{([\w.]+)(?:}|:)', fmt)
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
                warnings.warn("タイトル文字列に含まれる'{}'が，"
                              "dfに含まれないか，単一ではない．".format(key))
                dict_[key] = 'XXX'
        return fmt.format(**dict_)


def Excel(visible=True):
    if len(xw.apps) == 0:
        return xw.App(visible=visible)
    elif visible is xw.apps.active.visible:
        return xw.apps.active
    else:
        return xw.App(visible=visible)


def columns_list(df, columns):
    """
    ':column' or '::column' 形式を通常のカラム名のリストに変換する．
    ':column'はcolumnを含める，'::column'はcolumnの一つ前まで.
    文字列の場合はリストにする．

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
            if column.startswith('::'):
                yield from columns_[:columns_.index(column[2:])]
            elif column.startswith(':'):
                yield from columns_[:columns_.index(column[1:]) + 1]
            else:
                yield column

    return list(gen())


def delete_charts(sheet=None):
    """
    シート内のすべてのチャートを削除する．

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
        rcParams[f'axis.label.{key}'] = label
    if ticks:
        rcParams[f'axis.ticks.{key}'] = ticks
    if format:
        rcParams[f'axis.format.{key}'] = format


def autofilter(list_object, *args, **field_criteria):
    """
    キーワード引数で指定される条件に応じてフィルタリングする
    キーワード引数のキーはカラム名，値は条件．条件は以下のものが指定できる．
       - list : 要素を指定する．
       - tuple : 値の範囲を指定する．
       - None : 設定されているフィルタをクリアする
       - 他 : 値の一致

    """
    for field, criteria in zip(args[::2], args[1::2]):
        field_criteria[field] = criteria

    filter_ = list_object.Range.AutoFilter
    operator = xw.constants.AutoFilterOperator
    columns = [column.Name for column in list_object.ListColumns]

    for field, criteria in field_criteria.items():
        field_index = columns.index(field) + 1
        if isinstance(criteria, list):
            criteria = list(map(str, criteria))
            filter_(Field=field_index, Criteria1=criteria,
                    Operator=operator.xlFilterValues)
        elif isinstance(criteria, tuple):
            filter_(Field=field_index, Criteria1=f'>={criteria[0]}',
                    Operator=operator.xlAnd, Criteria2=f'<={criteria[1]}')
        elif criteria is None:
            filter_(Field=field_index)
        else:
            filter_(Field=field_index, Criteria1=f'{criteria}')


def main():
    sheet = xw.sheets.active
    list_object = sheet.api.ListObjects('テーブル1')
    autofilter(list_object, TMR=(100, 150))
