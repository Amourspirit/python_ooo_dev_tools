from __future__ import annotations
from typing import TypedDict


class ExportPngT(TypedDict):
    """
    Calc Image Export Data for Png.
    """

    pixel_width: int
    """Pixel Width"""
    pixel_height: int
    """Pixel Height"""
    compression: int
    """Compression value from ``1`` to ``9``. Higher compression, smaller file size. Default is ``6``."""
    interlaced: bool
    """Interlaced"""
    translucent: bool
    """Translucent"""
    logical_width: int
    """Logical Width"""
    logical_height: int
    """Logical Height"""
