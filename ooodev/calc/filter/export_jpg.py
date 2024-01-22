from __future__ import annotations
from ooodev.draw.filter import export_jpg


class ExportJpgT(export_jpg.ExportJpgT):
    """
    Calc Image Export Data for Jpg.

    - color_mode: Color Mode.
    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - quality: Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    """

    pass


class ExportJpg(export_jpg.ExportJpg):
    pass
