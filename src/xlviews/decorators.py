from __future__ import annotations

from functools import wraps
from typing import TYPE_CHECKING

import xlwings as xw

if TYPE_CHECKING:
    from collections.abc import Callable


def wait_updating(func: Callable) -> Callable:
    """Wait for the processing to finish before updating the screen."""

    @wraps(func)
    def _func(*args, **kwargs):  # noqa: ANN202
        active = xw.apps.active

        if active:
            is_updating = active.screen_updating
            active.screen_updating = False
        else:
            is_updating = None

        try:
            result = func(*args, **kwargs)
        finally:
            if active:
                active.screen_updating = is_updating

        return result

    return _func


def api(func: Callable) -> Callable:
    """If the first argument is an xlwings object, cast it to a win32com object."""

    @wraps(func)
    def _func(obj: object, *args, **kwargs):  # noqa: ANN202
        if hasattr(obj, "api"):
            obj = obj.api  # type: ignore

        return func(obj, *args, **kwargs)

    return _func
