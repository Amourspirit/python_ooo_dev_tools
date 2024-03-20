from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_height import UnitAppFontHeight

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class AppFontSize:
    """
    A class to represent the application font size.

    This class is a container for the ``Width`` and ``Height`` application font size units.

    Note:
        The  ``Width`` and ``Height`` units are based on the application font size ratio.
        This means that the  ``Width`` and ``Height`` units may not be based on the same unit length.
    """

    def __init__(self, width: UnitAppFontWidth, height: UnitAppFontHeight) -> None:
        """
        Constructor

        Args:
            width (UnitAppFontWidth): The  ``Width`` value of the application font size.
            height (UnitAppFontHeight): The ``Height`` value of the application font size.
        """
        self._width = width
        self._height = height

    @classmethod
    def from_app_font(cls, x: float, y: float) -> AppFontSize:
        """
        Creates an instance from ``1/100th mm`` values.

        Args:
            x (float): The  ``Width`` value of the application font size.
            y (float): The ``Height`` value of the application font size.

        Returns:
            SizeMM100: An instance of SizeMM100.
        """
        return cls(UnitAppFontWidth(x), UnitAppFontHeight(y))

    @property
    def width(self) -> UnitAppFontWidth:
        """
        Gets/Sets the ``Width```` unit of the application font size.

        When setting can be a float or a ``UnitT``.
        """
        return self._width

    @width.setter
    def width(self, value: float | UnitT) -> None:
        """
        Set the  ``Width`` value of the application font size.

        Parameters:
            value (float, UnitAppFontWidth): The new  ``Width`` value of the application font size.
        """
        self._width = UnitAppFontWidth.from_unit_val(value)

    @property
    def height(self) -> UnitAppFontHeight:
        """
        Get/Sets the ``Height`` unit of the application font size.

        When setting can be a float or a ``UnitT``.
        """
        return self._height

    @height.setter
    def height(self, value: float | UnitT) -> None:
        """
        Set the ``Height`` value of the application font size.

        Parameters:
            value (float, UnitAppFontHeight): The new ``Height`` value of the application font size.
        """
        self._height = UnitAppFontHeight.from_unit_val(value)
