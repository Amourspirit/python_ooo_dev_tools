from __future__ import annotations
from typing import NamedTuple


class FillColorProps(NamedTuple):
    """Internal Properties"""

    color: str  # FillColor ( int )
    style: str  # FillStyle such as FillStyle.SOLID
    bg: str = ""  # FillBackground ( bool, not used in all classes)
