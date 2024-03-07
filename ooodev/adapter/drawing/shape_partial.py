from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.drawing import XShape

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import Point
    from com.sun.star.awt import Size
    from ooodev.utils.type_var import UnoInterface


class ShapePartial:
    """
    Partial Class XShape.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShape, interface: UnoInterface | None = XShape) -> None:
        """
        Constructor

        Args:
            component (XShape): UNO Component that implements ``com.sun.star.drawing.XShape`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShape``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XShape
    def get_position(self) -> Point:
        """
        Gets the position of the shape in ``100th mm`` units.

        Returns:
            Point: The position of the shape.
        """
        return self.__component.getPosition()

    def set_position(self, position: Point) -> None:
        """
        Sets the position of the shape in ``100th mm`` units.

        Args:
            position (Point): The position of the shape.
        """
        self.__component.setPosition(position)

    def get_size(self) -> Size:
        """
        Gets the size of the shape in ``100th mm`` units.

        Returns:
            Size: The size of the shape.
        """
        return self.__component.getSize()

    def set_size(self, size: Size) -> None:
        """
        Sets the size of the shape in ``100th mm`` units.

        Args:
            size (Size): The size of the shape.
        """
        self.__component.setSize(size)

    # endregion XShape
