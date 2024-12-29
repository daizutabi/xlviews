from itertools import product

import numpy as np
import pandas as pd

from xlviews.decorators import wait_updating
from xlviews.frame import SheetFrame
import xlviews as xv
from xlviews.style import set_font, set_alignment
from xlviews.config import rcParams


class DistFrame(SheetFrame):
    @wait_updating
    def __init__(self, parent, columns=None, dist='norm', by=None, gray=True,
                 autofit=True):
        """
        分布をプロットするためにシートフレームを作成する．
        Parameters
        ----------
        dist : str or dict
        """
        self.by = parent.columns_list(by)
        self.dist_columns = columns if columns else parent.value_columns
        if isinstance(self.dist_columns, str):
            self.dist_columns = [self.dist_columns]
        if isinstance(dist, str):
            self.dist_func = {column: dist for column in self.dist_columns}
        else:
            self.dist_func = dist.copy()
            for column in self.dist_columns:
                if column not in self.dist_func:
                    self.dist_func[column] = 'norm'

        self.cell = parent.get_child_cell()

        # tableに変換する必要があるのはなぜ？
        if not parent.is_table:
            try:
                parent.astable()
            except Exception:
                pass
        self.parent = parent

        super().__init__(data=self.dummy_dataframe(), parent=self.parent,
                         index=self.by is not None, style=False)
        if self.by:
            self.link_to_index()
        grouped = self.groupby(self.by)
        for column in self.dist_columns:
            dist = self.dist_func[column]
            parent_column = self.parent.range(column).column
            column = self.range(column + '_n').column
            for key, row in grouped.items():
                if len(row) != 1:
                    raise ValueError('連続する行のみby可能')
                start = row[0][0]
                length = row[0][1] - start + 1
                parent_cell = self.parent.sheet.range(start, parent_column)
                cell = self.sheet.range(start, column)
                formula = counter(parent_cell)
                set_formula(cell, length, formula)
                formula = sorted_value(parent_cell, cell, length)
                set_formula(cell.offset(0, 1), length, formula)
                formula = sigma_value(cell, length, dist)
                set_formula(cell.offset(0, 2), length, formula)
        for column in self.dist_columns:
            parent_cell = self.parent.range(column)
            number_format = parent_cell.api.NumberFormatLocal
            self.set_number_format({f'{column}_v': number_format}, split=False)
            self.set_number_format({f'{column}_s': '0.00'}, split=False)
        self.set_style(autofit=autofit, gray=gray)
        self.const_values()

    def dummy_dataframe(self):
        columns = ['_'.join([column, name]) for column, name
                   in product(self.dist_columns, ['n', 'v', 's'])]
        array = np.zeros((len(self.parent), len(columns)))
        df = pd.DataFrame(array, columns=columns)
        if self.by:
            index = np.zeros((len(self.parent), len(self.by)))
            index = pd.DataFrame(index, columns=self.by)
            df = pd.concat([index, df], axis=1)  # type: pd.DataFrame
            df.set_index(self.by, inplace=True)
        return df

    def link_to_index(self):
        start = self.row + 1
        end = start + len(self) - 1
        for by in self.by:
            ref = self.parent.index(by)
            ref = self.parent.sheet.range(start, ref)
            ref = ref.get_address(row_absolute=False)
            formula = f'={ref}'
            to = self.index(by)
            range_ = self.sheet.range((start, to), (end, to))
            range_.value = formula

    def const_values(self):
        index = self.parent.index_columns
        array = np.zeros((len(index), 1))
        df = pd.DataFrame(array, columns=['value'], index=index)
        sf = xv.SheetFrame(data=df, head=self, gray=True, autofit=False)
        head = self.parent.cell.offset(-1, 0)
        tail = sf.cell.offset(1, 1)
        for k in range(len(index)):
            formula = '=' + head.offset(0, k).get_address()
            tail.offset(k, 0).value = formula

    # Plot関連
    def plot(self, x, label='auto', color=None, marker=None, axes=None,
             xlabel='auto', ylabel='auto', **kwargs):

        if ylabel == 'auto':
            dist = self.dist_func[x] if isinstance(x, str)\
                else self.dist_func[x[0]]
            ylabel = 'σ' if dist == 'norm' else 'ln(-ln(1-F))'
        plot = None
        if isinstance(x, str) and xlabel == 'auto':
            x_ = x.split('_')[0]
            xlabel = rcParams.get(f'axis.label.{x_}', x)
            if '_' in x and '[' in xlabel:
                xlabel = x + ' ' + xlabel[xlabel.index('['):]
        xs = [x] if isinstance(x, str) else x
        colors = color if isinstance(color, list) else [color] * len(xs)
        markers = marker if isinstance(marker, list) else [marker] * len(xs)
        for x_, color, marker in zip(xs, colors, markers):
            label_ = x_ if label == 'auto' and isinstance(x, list) else label
            plot = self._plot(x_, label=label_, axes=axes, color=color,
                              marker=marker, xlabel=xlabel, ylabel=ylabel,
                              **kwargs)
            axes = plot.axes
        return plot

    def _plot(self, x, **kwargs):
        plot = super().plot(f'{x}_v', f'{x}_s', yformat='0_ ', **kwargs)
        # if fit:
        #     sigma = 2 if fit is True else fit
        #     column = self.add_column_for_fit(x, sigma)
        #     kwargs['marker'] = None
        #     kwargs['line'] = None
        #     kwargs.pop('axes')
        #     kwargs.pop('label')
        #     plot_ = super().plot(f'{x}_v', column, axes=plot.axes,
        #                          label=None, **kwargs)
        #     for series in plot_.series_collection:
        #         trendline = series.Trendlines().Add()
        #         plot_.axes.labels.append('__trendline__')
        #         trendline.DisplayEquation = True
        #         # trendline.Forward = 10
        #         # trendline.Backward = 10
        #     # print(plot_.axes.labels)
        #     # print(plot_.axes.legend.LegendEntries())
        #     # print(plot_.legend)
        #     plot_.axes.set_legend(**plot_.legend)
        return plot

    def fit(self, x):
        pass

    def add_column_for_fit(self, x, sigma):
        """

        Parameters
        ----------
        x : str
            変数名
        sigma: int or float
            フィッティングに用いるσ値の範囲
        """
        column_ = f'{x}_sf'
        if column_ in self.columns:
            return column_
        self[column_] = 1
        column = self.index(column_)
        sigma_cell = self.sheet.range(self.row - 1, column)
        sigma_cell.value = sigma
        set_font(sigma_cell, size=8, bold=True, italic=True, color='green')
        set_alignment(sigma_cell, 'center')
        sigma = sigma_cell.get_address()
        row = self.cell.offset(self.columns_level).row
        column_ref = self.index(f'{x}_s')
        cell_ref = self.sheet.range(row, column_ref)
        cell_ref = cell_ref.get_address(row_absolute=False)
        cell = self.sheet.range(row, column)
        range_ = self.sheet.range(cell, cell.offset(len(self) - 1))
        range_.api.NumberFormatLocal = '0.00_ '
        formula = (f'=IF(AND({cell_ref}>=-{sigma},'
                   f'{cell_ref}<={sigma}),{cell_ref},NA())')
        range_.value = formula
        return column_


def counter(parent_cell):
    column = ':'.join([parent_cell.get_address(),
                       parent_cell.get_address(row_absolute=False)])
    return f'=AGGREGATE(3,1,{column})'
    # self.set_formula(formula, 0, offset, length)


def sorted_value(parent_cell, cell, length):
    end = parent_cell.offset(length - 1)
    column = ':'.join([parent_cell.get_address(), end.get_address()])

    small = cell.get_address(row_absolute=False)
    return f'=IF({small}>0,AGGREGATE(15,1,{column},{small}),NA())'
    # self.set_formula(formula, 1, offset, length)


def sigma_value(cell, length, dist):
    small = cell.get_address(row_absolute=False)
    end = cell.offset(length - 1)
    end = end.get_address()
    if dist == 'norm':
        return f'=IF({small}>0,NORM.S.INV({small}/({end}+1)),NA())'
    elif dist == 'weibull':
        return f'=IF({small}>0,LN(-LN(1-{small}/({end}+1))),NA())'
    else:
        raise ValueError('不明な分布', dist)
        # range_ = self.set_formula(formula, 2, offset, length)
        # range_.api.NumberFormatLocal = '0.00_ '


def set_formula(cell, length, formula):
    end = cell.offset(length - 1)
    cell.sheet.range(cell, end).value = formula


def main():
    import mtj
    import xlwings as xw

    xw.apps.add()
    book = xw.books[0]
    sheet = book.sheets[0]

    directory = mtj.get_directory('remote', 'Data')
    run = mtj.get_paths_dataframe(directory, 'SL1050-01', recipe='HR')
    series = run.iloc[0]
    path = mtj.get_path(directory, series)
    with mtj.data(path) as data:
        data.merge_device()
        sf = data.sheetframe(sheet, 2, 2,
                             columns=['wafer', 'cad', 'sx', 'sy', 'Rmin',
                                      'Rmax', 'TMR'],
                             index=':sy', sort_index=True)
    sf.set_number_format(Rmin='0.0', TMR='0')
    sf.distframe(by=':cad')
    # df.grid(x='cad', left=0).map('plot', ['Rmin', 'Rmax'],
    #                              xticks=(0, 30, 10),
    #                              yticks=(-4, 4), marker='o',
    #                              xlabel='Rmin [kΩ]', ylabel='σ', alpha=0.9,
    #                              fit=2)


if __name__ == '__main__':
    main()
