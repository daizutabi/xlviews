from __future__ import annotations

from itertools import product
from typing import TYPE_CHECKING

import numpy as np
import pandas as pd
from pandas import DataFrame

from xlviews.config import rcParams
from xlviews.decorators import turn_off_screen_updating
from xlviews.range.style import set_alignment, set_font

from .sheet_frame import SheetFrame

if TYPE_CHECKING:
    from xlwings import Range


class HeatFrame(SheetFrame):
    @turn_off_screen_updating
    def __init__(
        self,
        *args,
        data: DataFrame,
        x: str,
        y: str,
        value: str,
        style: bool = True,
        autofit: bool = True,
        **kwargs,
    ) -> None:
        df = data.pivot_table(value, y, x, aggfunc=lambda x: x)
        df.index.name = None

        super().__init__(*args, data=df, index=True, style=False, **kwargs)
