from __future__ import annotations
from typing import TypedDict


class ExportJpgT(TypedDict):
    """
    Draw Image Export Data for Jpg.

    - color_mode: Color Mode.
    - pixel_width: Pixel Width.
    - pixel_height: Pixel Height.
    - quality: Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``.
    - logical_width: Logical Width - Usually that same as ``pixel_width``.
    - logical_height: Logical Height - Usually that same as ``pixel_height``.
    """

    color_mode: bool
    pixel_width: int
    pixel_height: int
    quality: int
    logical_width: int
    logical_height: int


class ExportJpg(dict):
    def __init__(
        self,
        color_mode: bool,
        pixel_width: int,
        pixel_height: int,
        quality: int,
        logical_width: int,
        logical_height: int,
    ):
        """
        Constructor.

        Args:
            color_mode (bool): Color Mode.
            pixel_width (int): Pixel Width.
            pixel_height (int): Pixel Height.
            quality (int): Compression value from ``1`` to ``100``. Higher Quality, larger file size. Default is ``75``.
            logical_width (int): Logical Width - Usually that same as ``pixel_width``.
            logical_height (int): Logical Height - Usually that same as ``pixel_height``.
        """
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
