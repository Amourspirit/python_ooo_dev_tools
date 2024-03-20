from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class AppFontPos:
    """
    A class to represent the application font position.

    This class is a container for the ``X`` and ``Y`` application font position units.

    Note:
        The ``X`` and ``Y`` units are based on the application font position ratio.
        This means that the ``X`` and ``Y`` units may not not  be based on the same unit length.
    """

    def __init__(self, x: UnitAppFontX, y: UnitAppFontY) -> None:
        """
        Constructor

        Args:
            x (UnitAppFontX): The X value of the application font position.
            y (UnitAppFontY): The Y value of the application font position.
        """
        self._x = x
        self._y = y

    @classmethod
    def from_app_font(cls, x: float, y: float) -> AppFontPos:
        """
        Creates an instance from ``1/100th mm`` values.

        Args:
            x (float): The X value of the application font position.
            y (float): The Y value of the application font position.

        Returns:
            AppFontPos: An instance of AppFontPos.
        """
        return cls(UnitAppFontX(x), UnitAppFontY(y))

    @property
    def x(self) -> UnitAppFontX:
        """
        Gets/Sets the ``X`` unit of the application font position.

        When setting can be a float or a ``UnitT``.
        """
        return self._x

    @x.setter
    def x(self, value: float | UnitT) -> None:
        """
        Set the X value of the application font position.

        Parameters:
            value (float, UnitAppFontX): The new X value of the application font position.
        """
        self._x = UnitAppFontX.from_unit_val(value)

    @property
    def y(self) -> UnitAppFontY:
        """
        Get/Sets the ``Y`` unit of the application font position.

        When setting can be a float or a ``UnitT``.
        """
        return self._y

    @y.setter
    def y(self, value: float | UnitT) -> None:
        """
        Set the Y value of the application font position.

        Parameters:
            value (float, UnitAppFontY): The new Y value of the application font position.
        """
        self._y = UnitAppFontY.from_unit_val(value)
