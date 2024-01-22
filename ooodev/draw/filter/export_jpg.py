from __future__ import annotations
from typing import TypedDict


class ExportJpgT(TypedDict):
    """
    Draw Image Export Data for Jpg.

    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    - quality: Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``.
    - color_mode: Color Mode.
    """

    pixel_width: int
    pixel_height: int
    logical_width: int
    logical_height: int
    quality: int
    color_mode: bool


class ExportJpg(dict):
    def __init__(
        self,
        pixel_width: int,
        pixel_height: int,
        logical_width: int = -1,
        logical_height: int = -1,
        quality: int = 75,
        color_mode: bool = True,
    ):
        """
        Constructor.

        Args:
            pixel_width (int): Pixel Width.
            pixel_height (int): Pixel Height.
            logical_width (int): Logical Width - Usually that same as ``pixel_width``.
            logical_height (int): Logical Height - Usually that same as ``pixel_height``.
            quality (int): Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``.
            color_mode (bool): Color Mode.
        """
        if logical_width < 0:
            logical_width = pixel_width
        if logical_height < 0:
            logical_height = pixel_height
        if quality < 1 or quality > 100:
            quality = 75
        super().__init__(
            color_mode=color_mode,
            pixel_width=pixel_width,
            pixel_height=pixel_height,
            quality=quality,
            logical_width=logical_width,
            logical_height=logical_height,
        )

    def to_filter_dict(self):
        """
        Convert to Filter Data Dictionary.
        """
        return {
            "ColorMode": int(not self["color_mode"]),
            "PixelWidth": self["pixel_width"],
            "PixelHeight": self["pixel_height"],
            "Quality": self["quality"],
            "LogicalWidth": self["logical_width"],
            "LogicalHeight": self["logical_height"],
        }

    @staticmethod
    def from_dict(dict: ExportJpgT) -> ExportJpg:
        return ExportJpg(
            color_mode=dict["color_mode"],
            pixel_width=dict["pixel_width"],
            pixel_height=dict["pixel_height"],
            quality=dict["quality"],
            logical_width=dict["logical_width"],
            logical_height=dict["logical_height"],
        )
