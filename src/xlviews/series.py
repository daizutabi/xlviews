from __future__ import annotations

from typing import TYPE_CHECKING

from xlwings import Range

from xlviews.range import reference
from xlviews.style import set_series_style

if TYPE_CHECKING:
    from typing import Any

    from xlwings import Chart, Sheet
    from xlwings._xlwindows import COMRetryObjectWrapper


class Series:
    api: COMRetryObjectWrapper
    label: str

    def __init__(
        self,
        chart: Chart,
        x: Any,
        y: Any | None = None,
        label: str | tuple[int, int] | Range = "",
        chart_type: int | None = None,
        sheet: Sheet | None = None,
    ) -> None:
        self.api = chart.api[1].SeriesCollection().NewSeries()
        self.label = label if isinstance(label, str) else reference(label, sheet)
        self.name = self.label

        if chart_type is not None:
            self.chart_type = chart_type

        if y is not None:
            self.x = x
            self.y = y

        else:
            self.y = x

    @property
    def name(self) -> str:
        return str(self.api.Name)

    @name.setter
    def name(self, name: str) -> None:
        self.api.Name = name

    @property
    def chart_type(self) -> int:
        return self.api.ChartType  # type: ignore

    @chart_type.setter
    def chart_type(self, chart_type: int) -> None:
        self.api.ChartType = chart_type

    @property
    def x(self) -> tuple:
        return self.api.XValues  # type: ignore

    @x.setter
    def x(self, x: Any) -> None:
        if isinstance(x, Range):
            self.api.XValues = x.api
        else:
            self.api.XValues = x

    @property
    def y(self) -> tuple:
        return self.api.Values  # type: ignore

    @y.setter
    def y(self, y: Any) -> None:
        if isinstance(y, Range):
            self.api.Values = y.api
        else:
            self.api.Values = y

    @property
    def values(self) -> Range:
        return Range(self.api.Values)

    def set(self, **kwargs) -> None:
        set_series_style(self.api, **kwargs)

    def delete(self) -> None:
        self.api.Delete()
