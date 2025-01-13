from xlviews.axes import Axes
from xlviews.common import open_or_create
from xlviews.distframe import DistFrame
from xlviews.grouper import Grouper
from xlviews.series import Series
from xlviews.sheetframe import SheetFrame
from xlviews.statsframe import StatsFrame
from xlviews.utils import constant

__all__ = [
    "Axes",
    "DistFrame",
    "Grouper",
    "Series",
    "SheetFrame",
    "StatsFrame",
    "constant",
    "open_or_create",
]

# import win32com.client

# import matplotlib as mpl
# mpl.use('Qt5Agg')

# from xlviews.axes import Axes, chart_position
# from xlviews.dist import DistFrame
# from xlviews.element import Plot, Scatter, Element
# from xlviews.formula import interp1d, match_index, linear_fit
# from xlviews.frame import SheetFrame
# from xlviews.grid import FacetGrid
# from xlviews.powerpoint.main import PowerPoint
# from xlviews.shotmap import ShotMap
# from xlviews.stats import StatsFrame
# from xlviews.style import color_palette, marker_palette
# from xlviews.utils import (add_validation, copy_chart, copy_range, get_chart,
#                            get_range, get_sheet, open_or_create, Excel,
#                            delete_charts, set_axis_dimension)

# # Support for COM objects we use.
# win32com.client.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}',
#                                       0, 1, 3, bForDemand=True)  # Excel 9
# win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}',
#                                       0, 2, 1, bForDemand=True)  # Office
