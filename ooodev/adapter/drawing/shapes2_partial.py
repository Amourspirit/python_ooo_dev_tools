from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.drawing import XShapes2

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.utils.type_var import UnoInterface


class Shapes2Partial:
    """
    Partial class for XShapes2 interface.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapes2, interface: UnoInterface | None = XShapes2) -> None:
        """
        Constructor

        Args:
            component (XShapes2): UNO Component that implements ``com.sun.star.drawing.XShapes2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapes2``.
        """
        self.__component = component

    # region XShapes2
    def add_bottom(self, shape: XShape) -> None:
        """
        Adds a shape to the bottom of the stack.

        Args:
            shape (XShape): The shape to add.
        """
        self.__component.addBottom(shape)

    def add_top(self, shape: XShape) -> None:
        """
        Adds a shape to the top of the stack.

        Args:
            shape (XShape): The shape to add.
        """
        self.__component.addTop(shape)

    # endregion XShapes2
