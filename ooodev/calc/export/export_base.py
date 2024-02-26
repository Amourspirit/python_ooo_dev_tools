from __future__ import annotations
from ooodev.units.unit_inch import UnitInch


class ExportBase:
    """Base Class used for exporting Calc Shapes in various ways."""

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
        width_in = UnitInch.from_mm100(width)
        height_in = UnitInch.from_mm100(height)

        dpi_width = round(resolution * width_in.value)
        dpi_height = round(resolution * height_in.value)
        return dpi_width, dpi_height
