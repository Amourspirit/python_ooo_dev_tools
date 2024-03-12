from __future__ import annotations
from typing import Any, Generic, TypeVar, Union, TYPE_CHECKING
from ooodev.utils.data_type.generic_size_pos import GenericSizePos

from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound=UnitT)
TNum = TypeVar(name="TNum", bound=Union[int, float])
_TNum = TypeVar(name="_TNum", bound=Union[int, float])


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
