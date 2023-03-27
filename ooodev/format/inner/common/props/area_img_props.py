from __future__ import annotations
from typing import NamedTuple


class AreaImgProps(NamedTuple):
    """Internal Properties"""

    name: str  # FillBitmapName
    style: str  # FillStyle
    mode: str  # FillBitmapMode
    point: str  # FillBitmapRectanglePoint
    bitmap: str  # FillBitmap
    offset_x: str  # FillBitmapOffsetX
    offset_y: str  # FillBitmapOffsetY
    pos_x: str  # FillBitmapPositionOffsetX
    pos_y: str  # FillBitmapPositionOffsetY
    size_x: str  # FillBitmapSizeX
    size_y: str  # FillBitmapSizeY
