from __future__ import annotations
from typing import NamedTuple


class AreaPatternProps(NamedTuple):
    style: str  # FillStyle
    name: str  # FillBitmapName
    tile: str  # FillBitmapTile
    stretch: str  # FillBitmapStretch
    bitmap: str  # FillBitmap
