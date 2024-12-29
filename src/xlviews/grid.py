"""
マトリックス上にオブジェクトを配置するために、
FacetGrid, PairGridクラスを定義する。
Seabornの機能をチャート等にも適用できるようにする。
"""

import pandas as pd

# from xlviews.decorators import wait_updating
# from xlviews.chart import chart_position
from xlviews.axes import chart_position

# from spin.pandas import select
from xlviews.config import rcParams


class GridBase:
    def __init__(
        self,
        data,
        sheet=None,
        cell=None,
        left=None,
        top=None,
        width=None,
        height=None,
        row=None,
        column=None,
        space=0,
        callback=None,
    ):
        """
        GridBaseを生成する。

        Parameters
        ----------
        data : SheetFrame or DataFrame
            グリッド上に配置する元データ
        left, top : int, optional
            グリッドの開始位置
        width, height : int, optional
            オブジェクトのサイズ
        row, column : int, optional
            グリッドの開始位置をセルで指定する。
            この指定を行うと、シートフレームをグリッド上に配置する。
            mapで関数を呼び出すとき、func(sheet, row, column, ...)
        space : int or list, optional
            グリッド間の間隔
        """
        if cell is not None:
            sheet, row, column = cell.sheet, cell.row, cell.column
        if isinstance(space, int):
            space = space, space
        self.xspace, self.yspace = space
        self.data = data
        self.sheet = sheet if sheet else data.sheet

        if row and column:
            self.is_cellwise = True
            self.top, self.left = row, column
        else:
            self.is_cellwise = False
            self.left, self.top = chart_position(self.sheet, left, top)
            self.width = width or int(rcParams["chart.width"])
            self.height = height or int(rcParams["chart.height"])

        self.callback = callback
        self.elements = []
        self.element_dict = {}
        self.title = None

    def get_data(self):
        if hasattr(self.data, "columns_names"):
            if self.data.columns_names is None:
                data = self.data
            else:
                data = pd.DataFrame(
                    self.data.value_columns,
                    columns=self.data.columns_names,
                )
        else:
            data = self.data  # DataFrame or DataSet

        return data

    def __getattr__(self, item):
        def func(*args, **kwargs):
            return self.map(item, *args, **kwargs)

        return func

    # def columns_list(self, columns):
    #     if hasattr(self.data, 'columns_list'):
    #         return self.data.columns_list(x) if x else None
    #     else:
    #         return [x] if isinstance(x, str) else x


class FacetGrid(GridBase):
    def __init__(
        self,
        data,
        x=None,
        y=None,
        sheet=None,
        left=None,
        top=None,
        width=None,
        height=None,
        cell=None,
        row=None,
        column=None,
        space=0,
        callback=None,
    ):
        """
        FacetGridを生成する。

        Parameters
        ----------
        data : SheetFrame or DataFrame
            グリッド上に配置する元データ
        sheet
        left, top : int, optional
            グリッドの開始位置
        width, height : int, optional
            オブジェクトのサイズ
        row, column : int, optional
            グリッドの開始位置をセルで指定する。
            この指定を行うと、シートフレームをグリッド上に配置する。
            mapで関数を呼び出すとき、func(sheet, row, column, ...)
        x, y : str, optional
            列/行方向のカラム名
        space : int or list, optional
            グリッド間の間隔
        """
        super().__init__(
            data,
            sheet=sheet,
            cell=cell,
            left=left,
            top=top,
            width=width,
            height=height,
            row=row,
            column=column,
            space=space,
            callback=callback,
        )
        # if hasattr(data, 'columns_list'):
        #     self.x = data.columns_list(x) if x else None
        #     self.y = data.columns_list(y) if y else None
        # else:
        #     self.x = [x] if isinstance(x, str) else x
        #     self.y = [y] if isinstance(y, str) else y
        self.x = data.columns_list(x) if x else None
        self.y = data.columns_list(y) if y else None

        self.title = (self.x or []) + (self.y or [])

    def map(self, plot, *args, sel=None, by=None, title="auto", **kwargs):
        """
        FacetGridの各データに関数funcを適用する。
        funcの第一引数にはデータが渡される。
        """
        data = self.get_data()

        if self.y:
            row_df = data[self.y].drop_duplicates()
        else:
            row_df = pd.DataFrame([[0]], columns=["__dummy__"])
        if self.x:
            col_df = data[self.x].drop_duplicates()
        else:
            col_df = pd.DataFrame([[0]], columns=["__dummy__"])

        sel = sel if sel else {}
        if by is None:
            by = []
        elif not isinstance(by, list):
            by = [by]

        if isinstance(plot, str):
            plot = getattr(self.data, plot)
            data_kw = {}
        else:
            data_kw = {"data": self.data}

        if title == "auto":
            title = self.title

        self.elements = []
        self.element_dict = {}
        top = self.top
        for key_y, row_series in row_df.iterrows():
            if "__dummy__" in row_series:
                del row_series["__dummy__"]
            row_dict = dict(row_series)
            left = self.left
            row_elements = []
            element = None
            for key_x, col_series in col_df.iterrows():
                if "__dummy__" in col_series:
                    del col_series["__dummy__"]
                col_dict = dict(col_series)
                sel_ = dict(**sel, **col_dict, **row_dict)
                if self.is_cellwise:
                    element = plot(
                        self.sheet,
                        top,
                        left,
                        sel=sel_,
                        title=title,
                        **kwargs,
                        **data_kw,
                    )
                else:
                    element = plot(
                        *args,
                        sel=sel_,
                        by=by,
                        left=left,
                        top=top,
                        width=self.width,
                        height=self.height,
                        title=title,
                        sheet=self.sheet,
                        **kwargs,
                        **data_kw,
                    )
                    if element is None:
                        continue
                if self.is_cellwise:
                    left += element.width + 1
                else:
                    left += self.width
                if callable(self.callback):
                    self.callback(element, sel_)
                row_elements.append(element)
                key = []
                if self.x:
                    key += [col_dict[x_] for x_ in self.x]
                if self.y:
                    key += [row_dict[y_] for y_ in self.y]
                self.element_dict[tuple(key)] = element
            if self.is_cellwise:
                top += element.height + 1
            else:
                top += self.height
            self.elements.append(row_elements)
        return self.elements


class PairGrid(GridBase):
    def __init__(
        self,
        data,
        x_vars=None,
        y_vars=None,
        sheet=None,
        left=None,
        top=None,
        width=None,
        height=None,
        cell=None,
        row=None,
        column=None,
        space=0,
        callback=None,
    ):
        """
        PairGridを生成する。

        Parameters
        ----------
        data : SheetFrame or DataFrame
            グリッド上に配置する元データ
        left, top : int
            グリッドの開始位置
        width, height : int, optional
            オブジェクトのサイズ
        row, column : int, optional
            グリッドの開始位置をセルで指定する。
            この指定を行うと、シートフレームをグリッド上に配置する。
            mapで関数を呼び出すとき、func(sheet, row, column, ...)
        x_vars, y_vars : str, optional
            x, y のカラム名
        space : int or list, optional
            グリッド間の間隔
        """
        super().__init__(
            data,
            sheet=sheet,
            cell=cell,
            left=left,
            top=top,
            width=width,
            height=height,
            row=row,
            column=column,
            space=space,
            callback=callback,
        )
        if hasattr(data, "columns_list"):
            self.x_vars = data.columns_list(x_vars) if x_vars else None
            self.y_vars = data.columns_list(y_vars) if y_vars else None
        else:
            self.x_vars = [x_vars] if isinstance(x_vars, str) else x_vars
            self.y_vars = [y_vars] if isinstance(y_vars, str) else y_vars

    def map(
        self,
        plot,
        *args,
        sel=None,
        by=None,
        title="auto",
        transpose=False,
        **kwargs,
    ):
        """
        PairGridの各データに関数funcを適用する。
        funcの第一引数にはデータが渡される。
        """
        data = self.get_data()

        if self.y:
            row_df = data[self.y].drop_duplicates()
        else:
            row_df = pd.DataFrame([[0]], columns=["__dummy__"])
        if self.x:
            col_df = data[self.x].drop_duplicates()
        else:
            col_df = pd.DataFrame([[0]], columns=["__dummy__"])

        sel = sel if sel else {}
        if by is None:
            by = []
        elif not isinstance(by, list):
            by = [by]

        if isinstance(plot, str):
            plot = getattr(self.data, plot)
            data_kw = {}
        else:
            data_kw = {"data": self.data}

        if title == "auto":
            title = self.title

        self.elements = []
        top = self.top
        for key_y, row_series in row_df.iterrows():
            if "__dummy__" in row_series:
                del row_series["__dummy__"]
            row_dict = dict(row_series)
            left = self.left
            row_elements = []
            element = None
            for key_x, col_series in col_df.iterrows():
                if "__dummy__" in col_series:
                    del col_series["__dummy__"]
                col_dict = dict(col_series)
                sel_ = dict(**sel, **col_dict, **row_dict)
                if self.is_cellwise:
                    element = plot(
                        self.sheet,
                        top,
                        left,
                        sel=sel_,
                        title=title,
                        **kwargs,
                        **data_kw,
                    )
                else:
                    element = plot(
                        *args,
                        sel=sel_,
                        by=by,
                        left=left,
                        top=top,
                        width=self.width,
                        height=self.height,
                        title=title,
                        **kwargs,
                        **data_kw,
                    )
                    if element is None:
                        continue
                if self.is_cellwise:
                    left += element.width + 1
                else:
                    left += self.width
                if callable(self.callback):
                    self.callback(element, sel_)
                row_elements.append(element)
            if self.is_cellwise:
                top += element.height + 1
            else:
                top += self.height
            self.elements.append(row_elements)
        return self.elements


def main():
    # import xlwings as xw
    import xlviews as xv

    sf = xv.SheetFrame(2, 2, index_level=3, style=False)
    print(sf)


if __name__ == "__main__":
    main()
