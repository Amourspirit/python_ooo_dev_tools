from __future__ import annotations
from typing import Generic, TypeVar, Union
from ooodev.units.unit_obj import UnitT  # do not import from ooodev.unit or will cause circular import.
from ooodev.utils.data_type.generic_size import GenericSize

_T = TypeVar("_T", bound=UnitT)

TNum = TypeVar(name="TNum", bound=Union[int, float])
_TNum = TypeVar(name="_TNum", bound=Union[int, float])

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

    def get_size(self) -> GenericSize[TNum]:
        """Gets instance value as Size"""

        class Size(GenericSize[_TNum]):
            def __init__(self, width: _TNum, height: _TNum) -> None:
                super().__init__(width, height)

        return Size(self.width.value, self.height.value)  # type: ignore
