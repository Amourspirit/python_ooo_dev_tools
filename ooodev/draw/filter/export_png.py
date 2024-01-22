from __future__ import annotations
from typing import TypedDict


class ExportPngT(TypedDict):
    """
    Draw Image Export Data for Png.

    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - compression: Compression value from 1 to 9. Higher compression, smaller file size. Default is 6.
    - interlaced: Interlaced.
    - translucent: Translucent.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    """

    pixel_width: int
    pixel_height: int
    compression: int
    interlaced: bool
    translucent: bool
    logical_width: int
    logical_height: int


class ExportPng(dict):
    def __init__(
        self,
        pixel_width: int,
        pixel_height: int,
        compression: int,
        interlaced: bool,
        translucent: bool,
        logical_width: int,
        logical_height: int,
    ):
        """
        Constructor

        Args:
            pixel_width (int): Pixel Width.
            pixel_height (int): Pixel Height.
            compression (int): Compression value from 1 to 9. Higher compression, smaller file size. Default is 6.
            interlaced (bool): Interlaced.
            translucent (bool): Translucent.
            logical_width (int): Logical Width - Usually that same as ``pixel_width``.
            logical_height (int): Logical Height - Usually that same as ``pixel_height``.
        """
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
