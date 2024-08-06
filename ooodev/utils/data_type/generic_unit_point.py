from __future__ import annotations
from typing import Generic, TypeVar, Union
import uno
from com.sun.star.awt import Point

from ooodev.units.unit_convert import UnitLength
from ooodev.units import unit_factory
from ooodev.units.unit_obj import UnitT
from ooodev.utils.data_type.generic_point import GenericPoint

_T = TypeVar("_T", bound=UnitT)


# https://github.com/Amourspirit/python_ooo_dev_tools/issues/640
TNum = TypeVar("TNum", bound=Union[int, float])
_TNum = TypeVar("_TNum", bound=Union[int, float])

# example usage in: ooodev.form.controls.form_ctl_base.py


class GenericUnitPoint(Generic[_T, TNum]):
    """
    Point x and y.

    .. versionadded:: 0.17.3
    """

    def __init__(self, x: _T, y: _T) -> None:
        """
        Constructor

        Args:
            x (UnitT): Specifies width
            y (UnitT): Specifies height
        Returns:
            None:
        """
        self._x = x
        self._y = y

    # region Properties
    @property
    def x(self) -> _T:
        """Gets/Sets width"""
        return self._x

    @x.setter
    def x(self, value: _T):
        self._x = value

    @property
    def y(self) -> _T:
        """Gets/Sets height"""
        return self._y

    @y.setter
    def y(self, value: _T):
        self._y = value

    # endregion Properties

    def convert_to(self, unit_length: UnitLength) -> GenericUnitPoint[UnitT, Union[int, float]]:
        """
        Converts current values to specified unit length.

        Args:
            unit_length (UnitLength): Unit length to convert to.

        Returns:
            GenericUnitSize[UnitT, Union[int, float]]: Converted Units.
        """
        current_unit = self.x.get_unit_length()
        if current_unit == unit_length:
            return GenericUnitPoint(self.x, self.y)
        x = unit_factory.get_unit(unit_length, self.x.convert_to(unit_length))
        y = unit_factory.get_unit(unit_length, self.y.convert_to(unit_length))
        return GenericUnitPoint(x, y)

    def get_uno_point(self) -> Point:
        """Gets instance value as uno Point"""
        return Point(int(self.x), int(self.y))

    def get_point(self) -> GenericPoint[TNum]:
        """Gets instance value as Size"""

        class Point(GenericPoint[_TNum]):
            def __init__(self, x: _TNum, y: _TNum) -> None:
                super().__init__(x, y)

        return Point(self.x.value, self.y.value)  # type: ignore
