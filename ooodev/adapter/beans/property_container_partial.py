from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertyContainer

# com.sun.star.beans.PropertyAttribute
from ooo.dyn.beans.property_attribute import PropertyAttributeEnum

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PropertyContainerPartial:
    """
    Partial class for XPropertyContainer.
    """

    def __init__(self, component: XPropertyContainer, interface: UnoInterface | None = XPropertyContainer) -> None:
        """
        Constructor

        Args:
            component (XPropertyContainer): UNO Component that implements ``com.sun.star.container.XPropertyContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyContainer``.
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

    # region XPropertyContainer

    def add_property(self, name: str, attributes: PropertyAttributeEnum, value: Any) -> None:
        """
        Adds a property to the container.

        Args:
            name (str): The name of the property.
            attributes (PropertyAttributeEnum): The attributes of the property.
                Flags enum, this is a combination of ``com.sun.star.beans.PropertyAttribute``.
            value (Any): The value of the property.
        """
        self.__component.addProperty(name, attributes.value, value)

    def remove_property(self, name: str) -> None:
        """
        Removes a property from the container.

        Args:
            name (str): The name of the property.
        """
        self.__component.removeProperty(name)

    # endregion XPropertyContainer
