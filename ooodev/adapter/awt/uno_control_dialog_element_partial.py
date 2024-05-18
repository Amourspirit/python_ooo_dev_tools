from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialogElement  # Service
    from ooodev.units.unit_obj import UnitT
else:
    UnoControlDialogElement = Any


class UnoControlDialogElementPartial:
    """Partial class for UnoControlDialogElement."""

    def __init__(self, component: UnoControlDialogElement):
        """
        Constructor
        """
        self.__component = component

    # region Properties
    @property
    def height(self) -> UnitAppFontHeight:
        """
        Gets/Sets the height of the control.

        When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

        Returns:
            UnitAppFontHeight: Height of the control.
        """
        return UnitAppFontHeight(self.__component.Height)

    @height.setter
    def height(self, value: int | UnitT) -> None:
        val = UnitAppFontHeight.from_unit_val(value)
        self.__component.Height = int(val)

    @property
    def name(self) -> str:
        """
        Gets/Sets the name of the control.
        """
        return self.__component.Name

    @name.setter
    def name(self, value: str) -> None:
        self.__component.Name = value

    @property
    def x(self) -> UnitAppFontX:
        """
        Gets/Sets the horizontal position of the control.

        When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

        Returns:
            UnitAppFontX: Horizontal position of the control.
        """
        # the api is wrong, it should be int
        return UnitAppFontX(self.__component.PositionX)  # type: ignore

    @x.setter
    def x(self, value: int | UnitT) -> None:
        val = UnitAppFontX.from_unit_val(value)
        self.__component.PositionX = int(val)  # type: ignore

    @property
    def y(self) -> UnitAppFontY:
        """
        Gets/Sets the vertical position of the control.

        When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

        Returns:
            UnitAppFontY: Vertical position of the control.
        """
        # the api is wrong, it should be int
        return UnitAppFontY(self.__component.PositionY)  # type: ignore

    @y.setter
    def y(self, value: int | UnitT) -> None:
        val = UnitAppFontY.from_unit_val(value)
        self.__component.PositionY = int(val)  # type: ignore

    @property
    def step(self) -> int:
        """
        Gets/Sets the step of the control.
        """
        return self.__component.Step

    @step.setter
    def step(self, value: int) -> None:
        self.__component.Step = value

    @property
    def tab_index(self) -> int:
        """
        Gets/Sets the tab index of the control.
        """
        return self.__component.TabIndex

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        self.__component.TabIndex = value

    @property
    def tag(self) -> str:
        """
        Gets/Sets the tag of the control.
        """
        return self.__component.Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.__component.Tag = value

    @property
    def width(self) -> UnitAppFontWidth:
        """
        Gets/Sets the width of the control.

        When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

        Returns:
            UnitAppFontWidth: Width of the control.
        """
        return UnitAppFontWidth(self.__component.Width)

    @width.setter
    def width(self, value: int | UnitT) -> None:
        val = UnitAppFontWidth.from_unit_val(value)
        self.__component.Width = int(val)

    # endregion Properties
