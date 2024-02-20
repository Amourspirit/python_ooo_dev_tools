from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooodev.utils.partial.model_prop_partial import ModelPropPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialogElement  # Service
else:
    UnoControlDialogElement = Any


class UnoControlDialogElementPartial:
    """Partial class for UnoControlDialogElement. Must be used as a mixin that provides the ``ModelPropPartial``."""

    def __init__(self):
        """
        Constructor
        """
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlDialogElement

    # region Properties
    @property
    def height(self) -> int:
        """
        Gets/Sets the height of the control.
        """
        return self.model.Height

    @height.setter
    def height(self, value: int) -> None:
        self.model.Height = value

    @property
    def name(self) -> str:
        """
        Gets/Sets the name of the control.
        """
        return self.model.Name

    @name.setter
    def name(self, value: str) -> None:
        self.model.Name = value

    @property
    def position_x(self) -> int:
        """
        Gets/Sets the horizontal position of the control.
        """
        # the api is wrong, it should be int
        return self.model.PositionX  # type: ignore

    @position_x.setter
    def position_x(self, value: int) -> None:
        self.model.PositionX = value  # type: ignore

    @property
    def position_y(self) -> int:
        """
        Gets/Sets the vertical position of the control.
        """
        # the api is wrong, it should be int
        return self.model.PositionY  # type: ignore

    @position_y.setter
    def position_y(self, value: int) -> None:
        self.model.PositionY = value  # type: ignore

    @property
    def step(self) -> int:
        """
        Gets/Sets the step of the control.
        """
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_index(self) -> int:
        """
        Gets/Sets the tab index of the control.
        """
        return self.model.TabIndex

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        self.model.TabIndex = value

    @property
    def tag(self) -> str:
        """
        Gets/Sets the tag of the control.
        """
        return self.model.Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.model.Tag = value

    @property
    def width(self) -> int:
        """
        specifies the width of the control.
        """
        return self.model.Width

    @width.setter
    def width(self, value: int) -> None:
        self.model.Width = value

    # endregion Properties
