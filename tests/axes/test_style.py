import pytest
from xlwings import Sheet
from xlwings.constants import ChartType

from xlviews.axes import Axes


@pytest.fixture(scope="module")
def ax(sheet_module: Sheet):
    ct = ChartType.xlXYScatterLines
    ax = Axes(300, 10, chart_type=ct, sheet=sheet_module)
    x = sheet_module["B2:B11"]
    y = sheet_module["C2:C11"]
    z = sheet_module["D2:D11"]
    x.options(transpose=True).value = list(range(1, 11))
    y.options(transpose=True).value = list(range(10, 20))
    z.options(transpose=True).value = list(range(20, 30))
    sheet_module["C1"].value = "Y"
    sheet_module["D1"].value = "Z"
    ax.add_series(x, y, label=sheet_module["C1"])
    ax.add_series(x, z, label=(1, 4))
    return ax


def test_title_str(ax: Axes):
    ax.title = "a"
    assert ax.title == "a"


def test_title_range(ax: Axes, sheet_module: Sheet):
    cell = sheet_module["A1"]
    cell.value = "Title Range"
    ax.title = cell
    assert ax.title == "Title Range"


def test_title_tuple(ax: Axes, sheet_module: Sheet):
    sheet_module["A2"].value = "Title Tuple"
    ax.title = (2, 1)
    assert ax.title == "Title Tuple"


def test_title_none(ax: Axes):
    ax.title = None
    assert ax.title is None


def test_title_sytle(ax: Axes):
    ax.set_title("Style", name="Times", size=20, bold=True, italic=True)
    assert ax.title == "Style"
    chart_title = ax.chart.api[1].ChartTitle
    assert chart_title.Font.Name == "Times"
    assert chart_title.Font.Size == 20
    assert chart_title.Font.Bold
    assert chart_title.Font.Italic


def test_xlabel_str(ax: Axes):
    ax.xlabel = "a"
    assert ax.xlabel == "a"


def test_xlabel_range(ax: Axes, sheet_module: Sheet):
    cell = sheet_module["A3"]
    cell.value = "X Label Range"
    ax.xlabel = cell
    assert ax.xlabel == "X Label Range"


def test_xlabel_tuple(ax: Axes, sheet_module: Sheet):
    sheet_module["A4"].value = "X Label Tuple"
    ax.xlabel = (4, 1)
    assert ax.xlabel == "X Label Tuple"


def test_xlabel_none(ax: Axes):
    ax.xlabel = None
    assert ax.xlabel is None


def test_ylabel_str(ax: Axes):
    ax.ylabel = "a"
    assert ax.ylabel == "a"


def test_ylabel_range(ax: Axes, sheet_module: Sheet):
    cell = sheet_module["A5"]
    cell.value = "Y Label Range"
    ax.ylabel = cell
    assert ax.ylabel == "Y Label Range"


def test_ylabel_tuple(ax: Axes, sheet_module: Sheet):
    sheet_module["A6"].value = "Y Label Tuple"
    ax.ylabel = (6, 1)
    assert ax.ylabel == "Y Label Tuple"


def test_ylabel_none(ax: Axes):
    ax.ylabel = None
    assert ax.ylabel is None


def test_xticks(ax: Axes):
    ax.xticks = (5, 40, 10, 2)
    assert ax.xticks == (5, 40, 10, 2)


def test_yticks(ax: Axes):
    ax.yticks = (10, 40, 10, 2)
    assert ax.yticks == (10, 40, 10, 2)


def test_xticks_set(ax: Axes):
    ax.set_xticks(2, 5, 2, 1, gridlines=False)
    assert ax.xticks == (2, 5, 2, 1)


def test_yticks_set(ax: Axes):
    ax.set_yticks(2, 7, 2, 1, gridlines=False)
    assert ax.yticks == (2, 7, 2, 1)


def test_xtick_labels(ax: Axes):
    ax.set_xtick_labels("Arial", size=10, number_format="0.00")
    assert ax.xaxis.TickLabels.Font.Name == "Arial"
    assert ax.xaxis.TickLabels.Font.Size == 10
    assert ax.xaxis.TickLabels.NumberFormatLocal == "0.00"


def test_ytick_labels(ax: Axes):
    ax.set_ytick_labels("Times", size=14, number_format="0.0")
    assert ax.yaxis.TickLabels.Font.Name == "Times"
    assert ax.yaxis.TickLabels.Font.Size == 14
    assert ax.yaxis.TickLabels.NumberFormatLocal == "0.0"


@pytest.mark.parametrize("scale", ["log", "linear"])
def test_xscale(ax: Axes, scale: str):
    ax.xscale = scale
    assert ax.xscale == scale
    ax.xscale = "linear"


@pytest.mark.parametrize("scale", ["log", "linear"])
def test_yscale(ax: Axes, scale: str):
    ax.yscale = scale
    assert ax.yscale == scale
    ax.yscale = "linear"


def test_set(ax: Axes):
    ax.set(
        xlabel="1",
        ylabel="2",
        xticks=(1, 10, 4, 1),
        yticks=(2, 20, 4, 1),
        xscale="linear",
        yscale="linear",
    )
    assert ax.xlabel == "1"
    assert ax.ylabel == "2"
    assert ax.xticks == (1, 10, 4, 1)
    assert ax.yticks == (2, 20, 4, 1)
    assert ax.xscale == "linear"
    assert ax.yscale == "linear"


@pytest.mark.parametrize(
    ("loc", "left", "top"),
    [
        ((-1, -1), 25, 148),
        ((-1, 1), 25, 10),
        ((0, 0), 85, 80),
        ((1, 1), 145, 10),
        ((1, -1), 145, 148),
    ],
)
def test_legend_position(ax: Axes, loc, left, top):
    ax.set_legend(loc=loc)
    assert ax.chart.api[1].HasLegend
    # legend = ax2.chart.api[1].Legend
    # assert left - 3 < legend.Left < left + 3
    # assert top - 3 < legend.Top < top + 3


def test_tight_layout(ax: Axes):
    ax.xlabel = "x"
    ax.ylabel = "y"
    ax.title = "title"
    ax.tight_layout()
    assert 0 <= ax.chart.api[1].ChartTitle.Top <= 10

    ax.xlabel = None
    ax.tight_layout()


def test_style(ax: Axes):
    ax.set_style()
    assert ax.chart.api[1].PlotArea.Format.Line.Visible


def test_legend_none(sheet: Sheet):
    ct = ChartType.xlXYScatterLines
    ax = Axes(300, 10, chart_type=ct, sheet=sheet)
    x = sheet["B2:B11"]
    y = sheet["C2:C11"]
    z = sheet["D2:D11"]
    x.options(transpose=True).value = list(range(1, 11))
    y.options(transpose=True).value = list(range(10, 20))
    z.options(transpose=True).value = list(range(20, 30))
    ax.add_series(x, y)
    ax.add_series(x, z)
    ax.set_legend()
    assert ax.chart.api[1].HasLegend is False
