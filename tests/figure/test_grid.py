import pytest
from xlwings import Sheet

from xlviews.chart.axes import Axes
from xlviews.figure.grid import Grid, Series
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def seriesx(sheet_module: Sheet):
    ax = Axes(left=30, top=40, width=120, height=150, sheet=sheet_module)
    return Series(ax, 3, axis=0)


def test_series_len(seriesx: Series):
    assert len(seriesx) == 3


@pytest.mark.parametrize(
    ("k", "left", "top"),
    [(0, 30, 40), (1, 150, 40), (2, 270, 40)],
)
def test_seriesx_init(seriesx: Series, k: int, left, top):
    assert seriesx[k].chart.left == left
    assert seriesx[k].chart.top == top
    assert seriesx[:][k].chart.left == left
    assert seriesx[:][k].chart.top == top


@pytest.fixture(scope="module")
def seriesy(sheet_module: Sheet):
    ax = Axes(left=30, top=40, width=120, height=150, sheet=sheet_module)
    return Series(ax, 3, axis=1)


@pytest.mark.parametrize(
    ("k", "left", "top"),
    [(0, 30, 40), (1, 30, 190), (2, 30, 340)],
)
def test_seriesy_init(seriesy: Series, k: int, left, top):
    assert seriesy[k].chart.left == left
    assert seriesy[k].chart.top == top
    assert seriesy[:][k].chart.left == left
    assert seriesy[:][k].chart.top == top


def test_series_getitem_error(seriesx: Series):
    with pytest.raises(ValueError, match="Invalid key: abc"):
        seriesx["abc"]  # type: ignore


def test_series_iter(seriesx: Series):
    ax = list(seriesx)[-1]
    assert ax.chart.left == 270
    assert ax.chart.top == 40


def test_series_axis_error(seriesx):
    ax = seriesx[0]
    with pytest.raises(ValueError, match="Invalid axis: 2"):
        Series(ax, 1, axis=2)  # type: ignore


@pytest.fixture(scope="module")
def grid(sheet_module: Sheet):
    ax = Axes(left=30, top=40, width=120, height=150, sheet=sheet_module)
    return Grid(ax, 3, 4)


@pytest.mark.parametrize(
    ("r", "c", "left", "top"),
    [
        (0, 0, 30, 40),
        (0, 1, 150, 40),
        (0, 2, 270, 40),
        (0, 3, 390, 40),
        (1, 0, 30, 190),
        (1, 1, 150, 190),
        (1, 2, 270, 190),
        (1, 3, 390, 190),
        (2, 0, 30, 340),
        (2, 1, 150, 340),
        (2, 2, 270, 340),
        (2, 3, 390, 340),
    ],
)
def test_grid_init(grid: Grid, r: int, c: int, left, top):
    assert grid[r, c].chart.left == left
    assert grid[r, c].chart.top == top
    assert grid[r][c].chart.left == left
    assert grid[r][c].chart.top == top
    assert grid[:, c][r].chart.left == left
    assert grid[:, c][r].chart.top == top
    assert grid[r, :][c].chart.left == left
    assert grid[r, :][c].chart.top == top
    assert grid[:, :][r, c].chart.left == left


def test_grid_shape(grid: Grid):
    assert grid.shape == (3, 4)


def test_grid_len(grid: Grid):
    assert len(grid) == 3


def test_grid_shape_empty():
    grid = Grid([])
    assert grid.shape == (0, 0)


def test_grid_getitem_error(grid: Grid):
    with pytest.raises(ValueError, match="Invalid key: abc"):
        grid["abc"]  # type: ignore


def test_grid_iter(grid: Grid):
    assert next(iter(grid))[-1].chart.left == 390
    assert next(iter(grid))[-1].chart.top == 40
