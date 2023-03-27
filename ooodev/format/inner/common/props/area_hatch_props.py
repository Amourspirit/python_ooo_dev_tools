from __future__ import annotations
from typing import NamedTuple


class AreaHatchProps(NamedTuple):
    """Internal Properties"""

    color: str  # FillColor (int)
    style: str  # FillStyle such as FillStyle.HATCH
    bg: str  # FillBackground ( bool )
    hatch_prop: str  # FillHatch name for gradient struct property
