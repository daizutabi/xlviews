from __future__ import annotations

import re
from typing import TYPE_CHECKING

from win32com.client import gencache, selecttlb  # type: ignore

if TYPE_CHECKING:
    from collections.abc import Iterator

    from win32com.client.selecttlb import TypelibSpec


def iter_typelib_specs() -> Iterator[TypelibSpec]:
    pattern = r"Microsoft (Office|Excel|PowerPoint) \S+? Object Library"

    for tlb in selecttlb.EnumTlbs():
        if re.match(pattern, tlb.desc):
            yield tlb


def ensure_module(tlb: TypelibSpec) -> None:
    major = int(tlb.major, 16)
    minor = int(tlb.minor, 16)
    gencache.EnsureModule(tlb.clsid, tlb.lcid, major, minor)  # type: ignore


def ensure_modules() -> None:
    for tlb in iter_typelib_specs():
        ensure_module(tlb)
