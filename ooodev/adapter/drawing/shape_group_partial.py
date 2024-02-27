from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.drawing import XShapeGroup

from ooodev.adapter.drawing.shape_partial import ShapePartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ShapeGroupPartial(ShapePartial):
    """Partial class for XShapeGroup interface."""

    # Does no implement any methods.
    def __init__(self, component: XShapeGroup, interface: UnoInterface | None = XShapeGroup) -> None:
        """
        Constructor

        Args:
            component (XShapeGroup): UNO Component that implements ``com.sun.star.drawing.XShapeGroup`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapeGroup``.
        """
        ShapePartial.__init__(self, component, interface)
        self.__component = component

    # region XShapeGroup
    def enter_group(self) -> None:
        """
        Enters the group which enables the editing function for the parts of a grouped Shape.
        Then the parts can be edited instead of the group as a whole.
        This affects only the user interface. The behavior is not specified if this instance is not visible on any view. In this case it may or may not work.
        """
        self.__component.enterGroup()

    def leave_group(self) -> None:
        """
        leaves the group, which disables the editing function for the parts of a grouped Shape.

        Then only the group as a whole can be edited.

        This affects only the user interface. The behavior is not specified if this instance is not visible on any view. In this case it may or may not work.
        """
        self.__component.leaveGroup()

    # endregion XShapeGroup
