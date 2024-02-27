from __future__ import annotations
from typing import Generic, TypeVar, Union
from ooodev.units.unit_obj import UnitT
from ooodev.utils.data_type.generic_point import GenericPoint

_T = TypeVar("_T", bound=UnitT)

TNum = TypeVar(name="TNum", bound=Union[int, float])
_TNum = TypeVar(name="_TNum", bound=Union[int, float])

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

    def get_point(self) -> GenericPoint[TNum]:
        """Gets instance value as Size"""

        class Point(GenericPoint[_TNum]):
            def __init__(self, x: _TNum, y: _TNum) -> None:
                super().__init__(x, y)

        return Point(self.x.value, self.y.value)  # type: ignore
