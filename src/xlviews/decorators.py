from __future__ import annotations

from functools import wraps
from typing import TYPE_CHECKING

import xlwings as xw

if TYPE_CHECKING:
    from collections.abc import Callable


def turn_off_screen_updating(func: Callable) -> Callable:
    """Turn screen updating off to speed up your script."""

    @wraps(func)
    def _func(*args, **kwargs):  # noqa: ANN202
        if app := xw.apps.active:
            is_updating = app.screen_updating
            app.screen_updating = False

        try:
            return func(*args, **kwargs)
        finally:
            if app:
                app.screen_updating = is_updating

    return _func
