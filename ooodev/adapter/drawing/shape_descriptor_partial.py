from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.drawing import XShapeDescriptor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ShapeDescriptorPartial:
    """
    Class for managing IndexAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XShapeDescriptor, interface: UnoInterface | None = XShapeDescriptor) -> None:
        """
        Constructor

        Args:
            component (XShapeDescriptor): UNO Component that implements ``com.sun.star.drawing.XShapeDescriptor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XShapeDescriptor``.
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

    # region XShapeDescriptor
    def get_shape_type(self) -> str:
        """
        Gets the shape type.

        Returns:
            str: The  programmatic name of the shape type.
        """
        return self.__component.getShapeType()

    # endregion XShapeDescriptor
