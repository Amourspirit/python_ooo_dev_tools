from __future__ import annotations
from typing import Generic, TypeVar, Union, Tuple
import uno
from com.sun.star.awt import Point
from com.sun.star.awt import Size
from com.sun.star.awt import Rectangle


from ooodev.utils.data_type.generic_size_pos import GenericSizePos
from ooodev.units.unit_convert import UnitLength
from ooodev.units import unit_factory
from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound=UnitT)


# https://github.com/Amourspirit/python_ooo_dev_tools/issues/640
TNum = TypeVar("TNum", bound=Union[int, float])
_TNum = TypeVar("_TNum", bound=Union[int, float])


# class FloatSize(GenericSize[float]):
#     """Represents a size with positive values of Float."""

#     pass


# class IntSize(GenericSize[int]):
#     """Represents a size with positive values of Int."""

#     pass


class GenericUnitSizePos(Generic[_T, TNum]):
    """
    Size Width and Height

    .. versionadded:: 0.14.0
    """

    def __init__(self, x: _T, y: _T, width: _T, height: _T) -> None:
        """
        Constructor

        Args:
            width (UnitT): Specifies width
            height (UnitT): Specifies height
        Returns:
            None:
        """
        self._x = x
        self._y = y
        self._width = width
        self._height = height

    # region Properties
    @property
    def x(self) -> _T:
        """Gets/Sets x"""
        return self._x

    @x.setter
    def x(self, value: _T):
        self._x = value

    @property
    def y(self) -> _T:
        """Gets/Sets y"""
        return self._y

    @y.setter
    def y(self, value: _T):
        self._y = value

    @property
    def width(self) -> _T:
        """Gets/Sets width"""
        return self._width

    @width.setter
    def width(self, value: _T):
        self._width = value

    @property
    def height(self) -> _T:
        """Gets/Sets height"""
        return self._height

    @height.setter
    def height(self, value: _T):
        self._height = value

    def convert_to(self, unit_length: UnitLength) -> GenericUnitSizePos[UnitT, Union[int, float]]:
        """
        Converts current values to specified unit length.

        Args:
            unit_length (UnitLength): Unit length to convert to.

        Returns:
            GenericUnitSize[UnitT, Union[int, float]]: Converted Units.
        """
        current_unit = self.x.get_unit_length()
        if current_unit == unit_length:
            return GenericUnitSizePos(self.x, self.y, self.width, self.height)
        x = unit_factory.get_unit(unit_length, self.x.convert_to(unit_length))
        y = unit_factory.get_unit(unit_length, self.y.convert_to(unit_length))
        width = unit_factory.get_unit(unit_length, self.width.convert_to(unit_length))
        height = unit_factory.get_unit(unit_length, self.height.convert_to(unit_length))
        return GenericUnitSizePos(x, y, width, height)

    def get_uno_point_size(self) -> Tuple[Point, Size]:
        """
        Gets current values as Point and Size

        Returns:
            Tuple[Point, Size]: UNO Point and Size.
        """
        pos = Point()
        pos.X = int(self.x)
        pos.Y = int(self.y)
        size = Size()
        size.Width = int(self.width)
        size.Height = int(self.height)
        return pos, size

    def get_uno_rectangle(self) -> Rectangle:
        """Gets current values as Rectangle

        Returns:
            Rectangle: UNO Rectangle.
        """
        rect = Rectangle()
        rect.X = int(self.x)
        rect.Y = int(self.y)
        rect.Width = int(self.width)
        rect.Height = int(self.height)
        return rect

    # endregion Properties

    def get_size_pos(self) -> GenericSizePos[TNum]:
        """Gets instance value as Size"""

        class Size(GenericSizePos[_TNum]):
            def __init__(self, x: _TNum, y: _TNum, width: _TNum, height: _TNum) -> None:
                super().__init__(x, y, width, height)

        return Size(self.x.value, self.y.value, self.width.value, self.height.value)  # type: ignore


# class SizeMM(GenericUnitSizePos[UnitCM, float]):
#     pass


# size = SizeMM(UnitCM(2.0), UnitCM(3.0), UnitCM.from_cm(2), UnitCM.from_cm(3))
# size.height.value
# size.get_size_pos().y
