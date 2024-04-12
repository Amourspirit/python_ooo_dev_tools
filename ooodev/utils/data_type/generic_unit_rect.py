from __future__ import annotations
from typing import Generic, TypeVar, Union
import uno
from com.sun.star.awt import Rectangle

from ooodev.utils.data_type.generic_size_pos import GenericSizePos
from ooodev.units.unit_convert import UnitLength
from ooodev.units import unit_factory
from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound=UnitT)
TNum = TypeVar(name="TNum", bound=Union[int, float])
_TNum = TypeVar(name="_TNum", bound=Union[int, float])


class GenericUnitRect(Generic[_T, TNum]):
    """
    Rect X, Y, Width, Height

    .. versionadded:: 0.40.0
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

    # endregion Properties

    # region Methods
    def convert_to(self, unit_length: UnitLength) -> GenericUnitRect[UnitT, Union[int, float]]:
        """
        Converts current values to specified unit length.

        Args:
            unit_length (UnitLength): Unit length to convert to.

        Returns:
            GenericUnitSize[UnitT, Union[int, float]]: Converted Units.
        """
        current_unit = self.x.get_unit_length()
        if current_unit == unit_length:
            return GenericUnitRect(self.x, self.y, self.width, self.height)
        x = unit_factory.get_unit(unit_length, self.x.convert_to(unit_length))
        y = unit_factory.get_unit(unit_length, self.y.convert_to(unit_length))
        width = unit_factory.get_unit(unit_length, self.width.convert_to(unit_length))
        height = unit_factory.get_unit(unit_length, self.height.convert_to(unit_length))
        return GenericUnitRect(x, y, width, height)

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

    def get_size_pos(self) -> GenericSizePos[TNum]:
        """Gets instance value as Size"""

        class Rect(GenericSizePos[_TNum]):
            def __init__(self, x: _TNum, y: _TNum, width: _TNum, height: _TNum) -> None:
                super().__init__(x, y, width, height)

        return Rect(self.x.value, self.y.value, self.width.value, self.height.value)  # type: ignore

    # endregion Methods
