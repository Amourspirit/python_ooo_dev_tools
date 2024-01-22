from __future__ import annotations
from ooodev.utils import images_lo as mImage


class ExportBase:
    """Base Class used for exporting Draw Shapes in various ways."""

    def _get_dpi_width_height(self, width: int, height: int, resolution: int) -> tuple[int, int]:
        """
        Gets the DPI of the image.

        Args:
            width (int): Width in ``1/100 mm``.
            height (int): Height in ``1/100 mm``.
            resolution (int): Resolution in dpi.

        Returns:
            tuple[int, int]: x and y DPI.
        """
        return mImage.ImagesLo.get_dpi_width_height(width=width, height=height, resolution=resolution)
