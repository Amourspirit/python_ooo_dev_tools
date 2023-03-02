from __future__ import annotations
import math
from .size import Size
from ..unit_convert import UnitConvert
from ...proto.unit_obj import UnitObj


class SizeMM:
    """Size Width and Height in ``mm`` units."""

    def __init__(self, width: float | UnitObj, height: float | UnitObj) -> None:
        """
        Constructor

        Args:
            width (float, UnitObj, optional): Specifies width in ``mm`` units or :ref:`proto_unit_obj`.
            height (float, UnitObj, optional): Specifies height in ``mm`` units or :ref:`proto_unit_obj`.
        """
        self.width = width
        self.height = height

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

    def __eq__(self, other: object) -> bool:
        if isinstance(other, SizeMM):
            return math.isclose(self.width, other.width) and math.isclose(self.height, other.height)
        return NotImplemented

    # region Properties
    @property
    def width(self) -> float:
        """Gets/Sets width"""
        return self._width

    @width.setter
    def width(self, value: float | UnitObj):
        try:
            self._width = value.get_value_mm()
        except AttributeError:
            self._width = float(value)

    @property
    def height(self) -> float:
        """Gets/Sets height"""
        return self._height

    @height.setter
    def height(self, value: float | UnitObj):
        try:
            self._height = value.get_value_mm()
        except AttributeError:
            self._height = float(value)

    # endregion Properties
