from functools import wraps
from itertools import cycle

import pandas as pd

# import spin.pandas as spd
from xlviews.axes import Axes, chart_position
from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.style import palette
from xlviews.utils import format_label, label_func_from_list


def split_style(handle):
    """
    色指定などで、辞書を使うときの対応。
    color=('column', {'a': 'red', 'b': 'blue'})などを
    'column' と {('a',): 'red', 'b': 'blue'}に分割する。
    辞書のキーがタプルになっているのは、第一引数が最終的には
    リストになるため。
    color=('column', ['a', 'b', 'c'])の場合は、第2要素を順番に
    パレットから値を設定する。
    """
    if isinstance(handle, tuple):
        handle, style = handle
        if isinstance(style, dict):
            style = {
                key if isinstance(key, tuple) else (key,): value
                for key, value in style.items()
            }
        elif isinstance(style, list):
            style = [key if isinstance(key, tuple) else (key,) for key in style]
        return handle, style
    return handle, {}


def combine_handle(by, *handles, columns=None):
    """
    marker、colorなどをhandleをbyに追加する。
    なおかつ、リストでないものはリストにする。
    columnsを指定すると、marker, colorなどが直に形状や色を
    指定している場合を除外するために、columnsに含まれるものに限る。
    """

    def norm(x):
        if x is None:
            return []
        if not isinstance(x, list):
            return [x]
        return x[:]

    def none(x):
        return x if x else None

    by, *handles = [norm(x) for x in [by, *handles]]

    for handle in handles:
        for h in handle:
            if columns is not None and h not in columns:
                continue
            if h not in by:
                by.append(h)

    return [none(x) for x in [by, *handles]]


def get_label_name(label, by, key, **const_dict):
    """labelキーワード引数から凡例を作成する。"""
    if label and by:
        by_key = dict(zip(by, key, strict=False))
        for by_, key_ in const_dict.items():
            if by_ not in by:
                by_key[by_] = key_

        if isinstance(label, list):
            label = label_func_from_list(label)
        if callable(label):
            name = label(**by_key)
        else:
            name = label.format(**by_key)
    else:
        name = label
    if name == "auto":
        name = None
    return name


def autohandle(func):
    """
    マーカーや色のキーワード引数を処理する。
    この関数内で指定する任意のキーワード引数がhanldeになれる。
    これらのキーワード引数は、
      1. 'blue', 'o', など値を直に指定する。
      2. 'cad', 'pulse'などのようにカラム名で系列を分ける。
      3. ['cad', 'pulse']のように複数カラムで系列を分ける。
      4. ('cad', {'Y60': 'blue', 'Y70': 'red'})のように指定したスタイルで
          系列を分ける。
    """

    @wraps(func)
    def _func(self, *args, by=None, marker=None, color=None, alpha=None, **kwargs):
        marker, marker_style = split_style(marker)
        color, color_style = split_style(color)
        alpha, alpha_style = split_style(alpha)
        self.set_style(marker=marker_style, color=color_style, alpha=alpha_style)

        # if self.data.columns_level == 1:
        if self.data.columns_names is None:
            columns = self.data.columns
        else:
            columns = self.data.columns_names

        by, marker, color, alpha = combine_handle(
            by,
            marker,
            color,
            alpha,
            columns=columns,
        )
        self.set_handle(marker=marker, color=color, alpha=alpha)

        return func(self, *args, by=by, **kwargs)

    return _func


def autolegend(func):
    """
    凡例の処理を行う。
    """

    @wraps(func)
    def _func(self, *args, by=None, label=None, legend="auto", **kwargs):
        if legend == "auto":
            self.legend = dict(position=(1, 1))
        else:
            self.legend = legend
        if by and label == "auto":
            # label = '_'.join(['{' + x + '}' for x in by])
            label = label_func_from_list(by)

        axes = func(self, *args, by=by, label=label, **kwargs)

        if axes is None:
            return None

        if self.legend is None:
            axes.set_legend(None)
        elif isinstance(self.legend, dict):
            axes.set_legend(**self.legend)
        return axes

    return _func


class Element:
    def __init__(self, data):
        self.data = data
        self.sheet = None
        self.left = self.top = self.width = self.height = None
        self.handle = None
        self.const_dict = None
        self.style = None
        self.set_style()
        self.set_const_dict()
        self.axes = None
        self.chart = None
        self.series_collection = []
        self.by = None
        self.by_key = []
        self.legend = None
        self.x = None
        self.y = None
        self.chart_type = None
        self.include_in_layout = False

    def set_output(
        self,
        sheet=None,
        left=None,
        top=None,
        width=None,
        height=None,
        row=None,
        column=None,
        include_in_layout=False,
    ):
        self.sheet = sheet if sheet else self.data.sheet
        if column:
            self.left = self.sheet.range(1, column).left
        else:
            self.left = left
        if row:
            self.top = self.sheet.range(row, 1).top
        else:
            self.top = top
        self.left, self.top = chart_position(self.sheet, self.left, self.top)
        self.width, self.height = width, height
        self.include_in_layout = include_in_layout

    def set_style(self, **style):
        self.style = style.copy()

    def set_handle(self, **handle):
        self.handle = handle.copy()

    def set_const_dict(self, **const_dict):
        self.const_dict = const_dict.copy() if const_dict else {}
        # テールを持つシートフレームの場合は、そのデータを定数辞書に加える。
        if self.data.tail:
            self.const_dict.update(dict(self.data.tail.data.iloc[:, 0]))

    def set_axis(
        self,
        xticks="auto",
        yticks="auto",
        xlabel="auto",
        ylabel="auto",
        xformat="auto",
        yformat="auto",
        xscale=None,
        yscale=None,
    ):
        axes = self.axes
        if xscale:
            axes.set_xscale(xscale)
        if yscale:
            axes.set_yscale(yscale)

        x = self.x
        label_x = x[0] if isinstance(x, tuple) else x
        x = label_x.split("_")[0]
        y = self.y
        label_y = y[0] if isinstance(y, tuple) else y
        y = label_y.split("_")[0]

        if xticks == "auto":
            xticks = rcParams.get(f"axis.ticks.{x}", None)
        if yticks == "auto":
            yticks = rcParams.get(f"axis.ticks.{y}", None)
        if xticks:
            axes.set_xticks(*xticks)
        if yticks:
            axes.set_yticks(*yticks)

        if xformat == "auto":
            xformat = rcParams.get(f"axis.format.{x}", None)
        if yformat == "auto":
            yformat = rcParams.get(f"axis.format.{y}", None)
        axes.set_xtick_labels(format=xformat)
        axes.set_ytick_labels(format=yformat)

        if xlabel:
            if xlabel == "auto":
                xlabel = rcParams.get(f"axis.label.{x}", label_x)
                if "_" in label_x and "[" in xlabel:
                    xlabel = label_x + " " + xlabel[xlabel.index("[") :]
            axes.set_xlabel(xlabel)
        if ylabel:
            if ylabel == "auto":
                ylabel = rcParams.get(f"axis.label.{y}", label_y)
                if "_" in label_y and "[" in ylabel:
                    ylabel = label_y + " " + ylabel[ylabel.index("[") :]
            axes.set_ylabel(ylabel)

    @turn_off_screen_updating
    @autohandle
    @autolegend
    def _plot(
        self,
        x,
        y,
        chart_type=None,
        by=None,
        label=None,
        sel=None,
        sheet=None,
        left=None,
        top=None,
        width=None,
        height=None,
        row=None,
        column=None,
        axes=None,
        title="auto",
        line=False,
        line_width=False,
        name="{title_}_{x_}_{y_}",
        include_in_layout=False,
        **kwargs,
    ):
        """
        プロット本体

        Parameters
        ----------
        x, y : str
            シートフレームのカラム名
        chart_type : str
            エクセルチャートの種類
        by : str, list of str
            系列を分けるシートフレームのカラム名
        label : str
            ラベル文字列 '{cad}'などでシートフレームのカラムを含めることが
            できる。
        sel : dict
            シートフレームを選別する。
        sheet : xlwings.Sheet, optional
        left, top, width, height : int, optional
            チャートの位置と大きさ
        row, column : int, optional
            チャートのセル位置
        axes : xlviews.axes.Axes, optional
            重ね書きするときに指定する。
        title : str, optional
            タイトル
        line : bool, optional
            線を書くかどうか
        line_width : int, bool, optional
            線の太さ
        name : str
            チャートの名前
        data : xlviews.SheetFrame
            追加のデータをプロットするときのシートフレーム

        Returns
        -------
        xlviews.axes.Axes
        """
        if title == "auto" and hasattr(self.data, "title"):
            title = self.data.title

        if axes or self.axes:
            if axes:
                self.axes = axes
                self.chart = axes.chart
            is_new_axes = False
        else:
            self.set_output(
                sheet=sheet,
                left=left,
                top=top,
                width=width,
                height=height,
                row=row,
                column=column,
                include_in_layout=include_in_layout,
            )
            is_new_axes = True

        data = self.data
        self.x = x
        self.y = y
        self.chart_type = chart_type

        if isinstance(sel, dict):
            sel = data.select(**sel)
        else:
            sel = None
        grouped = data.groupby(by, sel)

        columns = data.columns
        if (
            data.columns_level == 1
            and (x not in columns or y not in columns)
            and not isinstance(x, tuple)
            and not isinstance(y, tuple)
        ):
            # Wide-format のデータフレーム。
            # 横方向にx軸、y軸をとる。
            by = data.index_columns
            if x in data.wide_columns and y in data.wide_columns:
                header = "both"
                index_x = data.index(x)
                index_y = data.index(y)
            elif x in data.wide_columns:
                header = "x"
                index = data.index(x)
            elif y in data.wide_columns:
                header = "y"
                index = data.index(y)
            else:
                raise ValueError("不明なカラム名", x, y)

            if sel is None:
                sel = [True] * len(data)
            if label == "auto":
                df = data[by][sel]
                df = spd.drop_const(df)
                label = list(df.columns)
                label = label_func_from_list(label)
            start = data.cell.row + 1
            column = data.cell.column
            for k in range(len(data)):
                row = start + k
                if sel[k]:
                    key = data.sheet.range(
                        (row, column),
                        (row, column + data.index_level),
                    ).value
                    series_name = get_label_name(label, by, key, **self.const_dict)
                    by_key = dict(zip(by, key, strict=False))
                    if header == "both":
                        # XValues
                        index = data.sheet.range((row, index_x[0]), (row, index_x[-1]))
                        # Values
                        columns = data.sheet.range(
                            (row, index_y[0]),
                            (row, index_y[-1]),
                        )
                    elif header == "x":
                        columns = [data.cell.row, row]
                    else:
                        columns = [row, data.cell.row]
                    self.add_series(index, columns, series_name, axis=0)
                    self.by_key.append(by_key)
        elif data.columns_names is None:
            # 通常のデータフレーム
            columns = data.index([x, y])
            for key, index in sorted(grouped.items()):
                if by is None:
                    by_key = None
                else:
                    by_key = dict(zip(by, key, strict=False))
                series_name = get_label_name(label, by, key, **self.const_dict)
                self.add_series(index, columns, series_name)
                self.by_key.append(by_key)
        else:
            # 波形など、unstackしたデータフレーム
            start = data.row + data.columns_level
            index = [start, start + len(data) - 1]

            # 軸項目はカラムの最後
            items = [column[-1] for column in data.columns]
            offset = data.column
            xycols = []
            for key, index_ in grouped.items():
                series_name = get_label_name(label, by, key, **self.const_dict)
                if by is None:
                    by_key = None
                else:
                    by_key = dict(zip(by, key, strict=False))

                # FIXME : もっと簡単にかけないか？
                for cols in index_:
                    xcol = ycol = None
                    for col in range(cols[0], cols[1] + 1):
                        item = items[col - offset]
                        if x == item:
                            if xcol is None:
                                xcol = col
                            else:
                                raise ValueError("重複したカラムあり。byを指定")
                        if y == item:
                            if ycol is None:
                                ycol = col
                            else:
                                raise ValueError("重複したカラムあり。byを指定")
                    if xcol and ycol:
                        xycols.extend([xcol - offset, ycol - offset])
                        self.add_series(index, [xcol, ycol], series_name)
                        self.by_key.append(by_key)
            # カラムがマルチインデックスの場合には、タイトル文字列用に
            # カラムインデックスからデータフレームを作成する。
            if title or name:
                data = pd.DataFrame(data.columns, columns=data.columns_names)
                data = data.iloc[xycols]
                sel = None

        if not self.axes:
            return None

        self.by = by
        self.apply_style(line=line, line_width=line_width)

        if is_new_axes:
            self.set_axis(**kwargs)
            if title:
                if isinstance(title, list):
                    title = label_func_from_list(title)
                if isinstance(title, str) or callable(title):
                    title = format_label(data, title, sel, default=self.const_dict)
                else:
                    name = None
                if title != "auto":
                    self.axes.set_title(title)
                else:
                    title = ""
            title = title if title else ""
            if name:
                name = name.replace("?", title)
                name = name.replace("{title_}", title)
                name = name.replace("{x_}", x[0] if isinstance(x, tuple) else x)
                name = name.replace("{y_}", y[0] if isinstance(y, tuple) else y)
                name = format_label(data, name, sel, default=self.const_dict)
                self.axes.chart.name = name
            self.axes.tight_layout()
            self.axes.set_style()
        return self.axes

    def add_axes(self):
        self.axes = Axes(
            sheet=self.sheet,
            left=self.left,
            top=self.top,
            width=self.width,
            height=self.height,
            include_in_layout=self.include_in_layout,
        )
        self.axes.set_chart_type(self.chart_type)
        self.chart = self.axes.chart

    def add_series(self, index, columns, name, axis=1):
        """
        シリーズを追加する。

        Parameters
        ----------
        index : list or xlwings.Range
            データインデックス
            See also: xlviews.utils.multirange
        columns : int or list or xlwings.Range
            intの場合、yの値のみ、listの場合(x, y)の値
        name : tuple or str
            tuple の場合、(row, col)
        axis : int
            データの方向

        Returns
        -------
        series : Series
            シリーズオブジェクト
        """
        if not self.axes:
            self.add_axes()
        sheet = self.data.sheet
        series = self.axes.add_series(
            index=index,
            columns=columns,
            name=name,
            axis=axis,
            sheet=sheet,
            chart_type=self.chart_type,
        )
        self.series_collection.append(series)
        return series

    @property
    def series_styles(self):
        """
        series_collectionに対応するスタイルのリストを返す。
        各要素はseriesに対応し、値は、{handle: style}の辞書である。
        """
        handles = self.handle.keys()
        styles = [self._series_styles(handle) for handle in handles]
        return [
            dict(zip(handles, style, strict=False))
            for style in zip(*styles, strict=False)
        ]

    def _series_styles(self, handle):
        """
        series_styleから呼ばれる内部関数で、handleごとのスタイルを生成する
        ジェネレータ
        """
        handles = self.handle[handle]
        if handles is None:
            yield from [None] * len(self.series_collection)
            return

        try:
            indexes = [self.by.index(handle) for handle in handles]
        except (ValueError, AttributeError):
            # hanldeに'red', 'cirlcle'などスタイルを直に描いている場合
            for _, handle in zip(self.series_collection, cycle(handles)):
                yield handle
            return

        defaults = palette(handle, len(self.series_collection))
        style = self.style[handle]
        if not isinstance(style, dict):
            style = dict(zip(style, palette(handle, len(style)), strict=False))
        for series, by_key, default in zip(
            self.series_collection,
            self.by_key,
            defaults,
            strict=False,
        ):
            key = tuple(by_key[self.by[index]] for index in indexes)
            if key not in style:
                style[key] = default
            else:
                defaults.insert(defaults.index(default), default)
            yield style[key]

    def apply_style(self, **kwargs):
        for series, style in zip(
            self.series_collection,
            self.series_styles,
            strict=False,
        ):
            set_series_style(series, **style, **kwargs)

    def plot(self, x, y, **kwargs):
        self._plot(x, y, chart_type="XYScatterLinesNoMarkers", **kwargs)

    def scatter(self, x, y, **kwargs):
        self._plot(x, y, chart_type="XYScatter", **kwargs)


class Plot(Element):
    def __init__(self, x, y, data=None, label="auto", **kwargs):
        super().__init__(data)
        self.plot(x, y, label=label, **kwargs)


class Scatter(Element):
    def __init__(self, x, y, data=None, marker="o", label="auto", **kwargs):
        super().__init__(data)
        self.scatter(x, y, marker=marker, label=label, **kwargs)


class Bar(Element):
    def __init__(self, x, y, data=None, label="auto", **kwargs):
        super().__init__(data)
        self.bar(x, y, label=label, **kwargs)

    @turn_off_screen_updating
    def bar(self, x, y, label="auto", stacked=False, **kwargs):
        if isinstance(x, list):
            xs, x = x, x[0]
        else:
            xs = None

        chart_type = "ColumnStacked" if stacked else "ColumnClustered"

        ys = y if isinstance(y, list) else [y]
        for y in ys:
            if label == "auto" and len(ys) > 1:
                label_ = y
            else:
                label_ = label
            self._plot(x, y, chart_type=chart_type, label=label_, **kwargs)

        if xs:
            sheet = self.data.sheet
            xvalues = sheet.range(
                self.data.range(xs[0])[0],
                self.data.range(xs[-1], -1)[-1],
            )
            self.series_collection[0].XValues = xvalues.api
            self.axes.set_xlabel(None)
