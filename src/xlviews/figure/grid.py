from __future__ import annotations

from typing import TYPE_CHECKING, overload

if TYPE_CHECKING:
    from collections.abc import Iterator
    from typing import Literal

    from xlviews.chart.axes import Axes


class Series:
    axes: list[Axes]

    def __init__(
        self,
        ax: Axes | list[Axes],
        n: int = 0,
        *,
        axis: Literal[0, 1] = 0,
    ) -> None:
        if isinstance(ax, list):
            self.axes = ax
            return

        if axis == 0:
            series = Grid(ax, 1, n)[0, :]
        elif axis == 1:
            series = Grid(ax, n, 1)[:, 0]
        else:
            msg = f"Invalid axis: {axis}"
            raise ValueError(msg)

        self.axes = list(series)

    @overload
    def __getitem__(self, key: int) -> Axes: ...

    @overload
    def __getitem__(self, key: slice) -> Series: ...

    def __getitem__(self, key: int | slice) -> Axes | Series:
        if isinstance(key, int):
            return self.axes[key]

        if isinstance(key, slice):
            return Series(self.axes[key])

        msg = f"Invalid key: {key}"
        raise ValueError(msg)

    def __len__(self) -> int:
        return len(self.axes)

    def __iter__(self) -> Iterator[Axes]:
        return iter(self.axes)


class Grid:
    axes: list[list[Axes]]

    def __init__(
        self,
        ax: Axes | list[list[Axes]],
        nrows: int = 0,
        ncols: int = 0,
    ) -> None:
        if isinstance(ax, list):
            self.axes = ax
            return

        left = ax.chart.left
        top = ax.chart.top
        width = ax.chart.width
        height = ax.chart.height

        axes = []
        for r in range(nrows):
            row = []
            for c in range(ncols):
                if r == 0 and c == 0:
                    row.append(ax)
                else:
                    new = ax.copy(left=left + c * width, top=top + r * height)
                    row.append(new)
            axes.append(row)
        self.axes = axes

    @property
    def shape(self) -> tuple[int, int]:
        if self.axes:
            return len(self.axes), len(self.axes[0])
        return 0, 0

    @overload
    def __getitem__(self, key: int) -> Series: ...

    @overload
    def __getitem__(self, key: tuple[int, int]) -> Axes: ...

    @overload
    def __getitem__(self, key: tuple[int, int]) -> Axes: ...

    @overload
    def __getitem__(self, key: tuple[slice, int]) -> Series: ...

    @overload
    def __getitem__(self, key: tuple[int, slice]) -> Series: ...

    @overload
    def __getitem__(self, key: tuple[slice, slice]) -> Grid: ...

    def __getitem__(self, key: int | tuple) -> Axes | Series | Grid:
        if isinstance(key, int):
            return Series(self.axes[key])

        if isinstance(key, tuple) and len(key) == 2:
            r, c = key
            if isinstance(r, int) and isinstance(c, int):
                return self.axes[r][c]

            if isinstance(r, slice) and isinstance(c, int):
                return Series([row[c] for row in self.axes[r]])

            if isinstance(r, int) and isinstance(c, slice):
                return Series(self.axes[r][c])

            if isinstance(r, slice) and isinstance(c, slice):
                return Grid([row[c] for row in self.axes[r]])

        msg = f"Invalid key: {key}"
        raise ValueError(msg)

    def __len__(self) -> int:
        return len(self.axes)

    def __iter__(self) -> Iterator[Series]:
        for row in self.axes:
            yield Series(row)
