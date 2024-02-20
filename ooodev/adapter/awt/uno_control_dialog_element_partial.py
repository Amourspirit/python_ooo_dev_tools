from __future__ import annotations
from typing import TYPE_CHECKING
import uno  # pylint: disable=unused-import

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialogElement  # Service


class UnoControlDialogElementPartial:
    """Partial class for UnoControlDialogElement."""

    def __init__(self, component: UnoControlDialogElement):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlDialogElement`` service.
        """
        # pylint: disable=unused-argument
        self.__component = component

    # region Properties
    @property
    def height(self) -> int:
        """
        Gets/Sets the height of the control.
        """
        return self.__component.Height

    @height.setter
    def height(self, value: int) -> None:
        self.__component.Height = value

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
    def position_x(self) -> int:
        """
        Gets/Sets the horizontal position of the control.
        """
        # the api is wrong, it should be int
        return self.__component.PositionX  # type: ignore

    @position_x.setter
    def position_x(self, value: int) -> None:
        self.__component.PositionX = value  # type: ignore

    @property
    def position_y(self) -> int:
        """
        Gets/Sets the vertical position of the control.
        """
        # the api is wrong, it should be int
        return self.__component.PositionY  # type: ignore

    @position_y.setter
    def position_y(self, value: int) -> None:
        self.__component.PositionY = value  # type: ignore

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
    def width(self) -> int:
        """
        specifies the width of the control.
        """
        return self.__component.Width

    @width.setter
    def width(self, value: int) -> None:
        self.__component.Width = value

    # endregion Properties
