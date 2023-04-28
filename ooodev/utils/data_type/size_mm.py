from __future__ import annotations
from .size import Size
from ooodev.units import UnitConvert
from ooodev.units import UnitObj
from ooodev.proto.size_obj import SizeObj


class SizeMM:
    """Size Width and Height in ``mm`` units."""

    def __init__(self, width: float | UnitObj, height: float | UnitObj) -> None:
        """
        Constructor

        Args:
            width (float, UnitObj): Specifies width in ``mm`` units or :ref:`proto_unit_obj`.
            height (float, UnitObj): Specifies height in ``mm`` units or :ref:`proto_unit_obj`.

        Returns:
            None:
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
    def from_size_mm100(size: SizeObj) -> SizeMM:
        """
        Gets instance from Size where Size is in ``1/100th mm`` units.

        Args:
            size (Size): Size in ``1/100th mm`` units

        Returns:
            SizeMM: Size in mm units
        """
        return SizeMM(width=UnitConvert.convert_mm100_mm(size.Width), height=UnitConvert.convert_mm100_mm(size.Height))

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

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, SizeMM):
            w1 = round(round(self.width, 2) * 100)
            w2 = round(round(oth.width, 2) * 100)
            wrng = range(w1 - 2, w1 + 3)  # +- 2
            if not w2 in wrng:
                return False
            h1 = round(round(self.height, 2) * 100)
            h2 = round(round(oth.height, 2) * 100)
            hrng = range(h1 - 2, h1 + 3)  # +- 2
            return h2 in hrng
            # return math.isclose(self.width, oth.width) and math.isclose(self.height, oth.height)
        return NotImplemented

    # region Properties
    @property
    def width(self) -> float:
        """Gets/Sets width"""
        return self._width

    @width.setter
    def width(self, value: float | UnitObj):
        try:
            self._width = round(value.get_value_mm(), 2)
        except AttributeError:
            self._width = round(float(value), 2)

    @property
    def height(self) -> float:
        """Gets/Sets height"""
        return self._height

    @height.setter
    def height(self, value: float | UnitObj):
        try:
            self._height = round(value.get_value_mm(), 2)
        except AttributeError:
            self._height = round(float(value), 2)

    # endregion Properties
