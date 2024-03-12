from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.drawing import XControlShape

from ooodev.adapter.drawing.shape_partial import ShapePartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.awt import XControlModel


class ControlShapePartial(ShapePartial):
    """
    Partial Class for XControlShape.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XControlShape, interface: UnoInterface | None = XControlShape) -> None:
        """
        Constructor

        Args:
            component (XControlShape): UNO Component that implements ``com.sun.star.drawing.XControlShape`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControlShape``.
        """
        ShapePartial.__init__(self, component, interface)
        self.__component = component

    # region XControlShape
    def get_control(self) -> XControlModel:
        """
        returns the control model of this Shape.
        """
        return self.__component.getControl()

    def set_control(self, control: XControlModel) -> None:
        """
        sets the control model for this Shape.
        """
        self.__component.setControl(control)

    # endregion XControlShape
