from __future__ import annotations
from typing import TypedDict


class ExportJpgT(TypedDict):
    """
    Calc Image Export Data for Jpg.
    """

    color_mode: bool
    """Color Mode"""
    pixel_width: int
    """Pixel Width"""
    pixel_height: int
    """Pixel Height"""
    quality: int
    """Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``."""
    logical_width: int
    """Logical Width"""
    logical_height: int
    """Logical Height"""
