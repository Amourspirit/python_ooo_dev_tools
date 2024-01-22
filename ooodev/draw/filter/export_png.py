from __future__ import annotations
from typing import TypedDict


class ExportPngT(TypedDict):
    """
    Draw Image Export Data for Png.

    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    - compression: Compression value from 1 to 9. Higher compression, smaller file size. Default is 6.
    - translucent: Translucent.
    - interlaced: Interlaced.
    """

    pixel_width: int
    pixel_height: int
    logical_width: int
    logical_height: int
    compression: int
    translucent: bool
    interlaced: bool


class ExportPng(dict):
    def __init__(
        self,
        pixel_width: int,
        pixel_height: int,
        logical_width: int = -1,
        logical_height: int = -1,
        compression: int = 6,
        translucent: bool = False,
        interlaced: bool = False,
    ):
        """
        Constructor

        Args:
            pixel_width (int): Pixel Width.
            pixel_height (int): Pixel Height.
            logical_width (int, optional): Logical Width - Usually that same as ``pixel_width``. Defaults to Pixel Width.
            logical_height (int, optional): Logical Height - Usually that same as ``pixel_height``. Defaults to Pixel Height.
            compression (int, optional): Compression value from 1 to 9. Higher compression, smaller file size. Default is ``6``.
            translucent (bool): Translucent. Defaults to ``False``.
            interlaced (bool): Interlaced. Defaults to ``False``.
        """
        if logical_width < 0:
            logical_width = pixel_width
        if logical_height < 0:
            logical_height = pixel_height
        if compression < 1 or compression > 9:
            compression = 6
        super().__init__(
            pixel_width=pixel_width,
            pixel_height=pixel_height,
            compression=compression,
            interlaced=interlaced,
            translucent=translucent,
            logical_width=logical_width,
            logical_height=logical_height,
        )

    def to_filter_dict(self):
        """
        Convert to Filter Data Dictionary.
        """
        return {
            "PixelWidth": self["pixel_width"],
            "PixelHeight": self["pixel_height"],
            "Compression": self["compression"],
            "Interlaced": int(self["interlaced"]),
            "Translucent": int(self["translucent"]),
            "LogicalWidth": self["logical_width"],
            "LogicalHeight": self["logical_height"],
        }

    @staticmethod
    def from_dict(dict: ExportPngT) -> ExportPng:
        return ExportPng(
            pixel_width=dict["pixel_width"],
            pixel_height=dict["pixel_height"],
            compression=dict["compression"],
            interlaced=dict["interlaced"],
            translucent=dict["translucent"],
            logical_width=dict["logical_width"],
            logical_height=dict["logical_height"],
        )
