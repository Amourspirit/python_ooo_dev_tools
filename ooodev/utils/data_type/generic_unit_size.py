from __future__ import annotations
from typing import Generic, TypeVar, Union
import uno
from com.sun.star.awt import Size

from ooodev.units.unit_convert import UnitLength
from ooodev.units import unit_factory
from ooodev.units.unit_obj import UnitT  # do not import from ooodev.unit or will cause circular import.
from ooodev.utils.data_type.generic_size import GenericSize

_T = TypeVar("_T", bound=UnitT)
# https://github.com/Amourspirit/python_ooo_dev_tools/issues/640
TNum = TypeVar("TNum", bound=Union[int, float])
_TNum = TypeVar("_TNum", bound=Union[int, float])


# example usage in: ooodev.form.controls.form_ctl_base.py


class GenericUnitSize(Generic[_T, TNum]):
    """
    Size Width and Height.

    .. versionadded:: 0.14.0
    """

    def __init__(self, width: _T, height: _T) -> None:
        """
        Constructor

        Args:
            width (UnitT): Specifies width
            height (UnitT): Specifies height
        Returns:
            None:
        """
        self._width = width
        self._height = height

    # region Properties
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

    def convert_to(self, unit_length: UnitLength) -> GenericUnitSize[UnitT, Union[int, float]]:
        """
        Converts current values to specified unit length.

        Args:
            unit_length (UnitLength): Unit length to convert to.

        Returns:
            GenericUnitSize[UnitT, Union[int, float]]: Converted Units.
        """
        current_unit = self.height.get_unit_length()
        if current_unit == unit_length:
            return GenericUnitSize(self.width, self.height)
        width = unit_factory.get_unit(unit_length, self.width.convert_to(unit_length))
        height = unit_factory.get_unit(unit_length, self.height.convert_to(unit_length))
        return GenericUnitSize(width, height)

    def get_uno_size(self) -> Size:
        """
        Gets current values as Size

        Returns:
            Size: UNO Size instance
        """
        size = Size()
        size.Width = int(self.width)
        size.Height = int(self.height)
        return size

    def get_size(self) -> GenericSize[TNum]:
        """Gets instance value as Size"""

        class Size(GenericSize[_TNum]):
            def __init__(self, width: _TNum, height: _TNum) -> None:
                super().__init__(width, height)

        return Size(self.width.value, self.height.value)  # type: ignore
