from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.drawing import XShapeGrouper
from com.sun.star.drawing import XShapeGroup
from com.sun.star.drawing import XShapes

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.utils.type_var import UnoInterface


class ShapeGrouperPartial:
    """
    Partial class for XShapeGrouper interface.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapeGrouper, interface: UnoInterface | None = XShapeGrouper) -> None:
        """
        Constructor

        Args:
            component (XShapeGrouper): UNO Component that implements ``com.sun.star.drawing.XShapeGrouper`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapeGrouper``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XShapeGrouper
    def group(self, shapes: XShapes) -> XShapeGroup:
        """
        Groups the given shapes.

        Args:
            shapes (XShapes): The shapes to group.

        Returns:
            XShapeGroup: The group shape.
        """
        return self.__component.group(shapes)

    def ungroup(self, shape: XShapeGroup) -> None:
        """
        Un-groups the given shape.

        Args:
            shape (XShape): The shape to ungroup.

        Returns:
            XShapes: The ungrouped shapes.
        """
        self.__component.ungroup(shape)
    # endregion XShapeGrouper
