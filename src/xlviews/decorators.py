from __future__ import annotations

from functools import wraps
from typing import TYPE_CHECKING, ParamSpec, TypeVar

import xlwings as xw

if TYPE_CHECKING:
    from collections.abc import Callable

P = ParamSpec("P")
R = TypeVar("R")


def turn_off_screen_updating(func: Callable[P, R]) -> Callable[P, R]:
    """Turn screen updating off to speed up your script."""

    @wraps(func)
    def _func(*args: P.args, **kwargs: P.kwargs) -> R:
        if app := xw.apps.active:
            is_updating = app.screen_updating
            app.screen_updating = False

        try:
            return func(*args, **kwargs)
        finally:
            if app:
                app.screen_updating = is_updating

    return _func
