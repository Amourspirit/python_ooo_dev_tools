from __future__ import annotations
from ooodev.draw.filter import export_png


class ExportPngT(export_png.ExportPngT):
    """
    Calc Image Export Data for Jpg.

    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - compression: Compression value from 1 to 9. Higher compression, smaller file size. Default is 6.
    - interlaced: Interlaced.
    - translucent: Translucent.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    """

    pass


class ExportPng(export_png.ExportPng):
    pass
