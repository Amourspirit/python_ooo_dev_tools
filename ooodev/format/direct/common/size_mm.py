from __future__ import annotations
from dataclasses import dataclass
from ....utils.data_type.size import Size
from ....utils.data_type.width_height_fraction import WidthHeightFraction
from ....utils.unit_convert import UnitConvert


@dataclass(frozen=True)
class SizeMM(WidthHeightFraction):
    """Size Width and Height in ``mm`` units."""

    def get_size_mm100(self) -> Size:
        """
        Gets instance converted to Size in ``1/100th mm`` units.

        Returns:
            Size: Size in ``mm`` units.
        """
        return Size(width=UnitConvert.convert_mm_mm100(self.width), height=UnitConvert.convert_mm_mm100(self.height))

    @staticmethod
    def from_size_mm100(size: Size) -> SizeMM:
        """
        Gets instance from Size where Size is in ``1/100th mm`` units.

        Args:
            size (Size): Size in ``1/100th mm`` units

        Returns:
            SizeMM: Size in mm units
        """
        return SizeMM(width=UnitConvert.convert_mm100_mm(size.width), height=UnitConvert.convert_mm100_mm(size.height))

    @staticmethod
    def from_mm100(width: int, height: int) -> SizeMM:
        """
        Gets instance from width and height where width and height are in ``1/100th mm`` units.

        Args:
            width (int): Width in ``1/100th mm`` units.
            height (int): Height in ``1/100th mm`` units.

        Returns:
            SizeMM: Size in mm units
        """
        return SizeMM(width=UnitConvert.convert_mm100_mm(width), height=UnitConvert.convert_mm100_mm(height))
