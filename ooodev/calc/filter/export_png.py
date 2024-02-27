from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.draw.filter import export_png

    class ExportPngT(export_png.ExportPngT):
        """
        Calc Image Export Data for Jpg.

        - pixel_width: Pixel Width.
        - pixel_height: Pixel Height.
        - logical_width: Logical Width - Usually that same as ``pixel_width``.
        - logical_height: Logical Height - Usually that same as ``pixel_height``.
        - compression: Compression value from 1 to 9. Higher compression, smaller file size. Default is 6.
        - translucent: Translucent.
        - interlaced: Interlaced.
        """

        pass

    class ExportPng(export_png.ExportPng):
        pass

else:
    ExportPngT = object
    ExportPng = object
