from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertyChangeListener
    from com.sun.star.beans import XVetoableChangeListener
    from com.sun.star.beans import XPropertySetInfo
    from ooodev.utils.type_var import UnoInterface


class PropertySetPartial:
    """
    Partial class for XPropertySet.
    """

    def __init__(self, component: XPropertySet, interface: UnoInterface | None = XPropertySet) -> None:
        """
        Constructor

        Args:
            component (XPropertySet): UNO Component that implements ``com.sun.star.beans.XPropertySet`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertySet``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertySet

    def add_property_change_listener(self, name: str, listener: XPropertyChangeListener) -> None:
        """
        Adds a listener for property changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be added.
        """
        self.__component.addPropertyChangeListener(name, listener)

    def add_vetoable_change_listener(self, name: str, listener: XVetoableChangeListener) -> None:
        """
        Adds a listener for vetoable changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be added.
        """
        self.__component.addVetoableChangeListener(name, listener)

    def get_property_set_info(self) -> XPropertySetInfo:
        """
        Returns the property set info.

        Returns:
            XPropertySetInfo: The property set info.
        """
        return self.__component.getPropertySetInfo()

    def get_property_value(self, name: str) -> Any:
        """
        Returns the value of a property.

        Args:
            name (str): The name of the property.

        Returns:
            Any: The value of the property.
        """
        return self.__component.getPropertyValue(name)

    def remove_property_change_listener(self, name: str, listener: XPropertyChangeListener) -> None:
        """
        Removes a listener for property changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be removed.
        """
        self.__component.removePropertyChangeListener(name, listener)

    def remove_vetoable_change_listener(self, name: str, listener: XVetoableChangeListener) -> None:
        """
        Removes a listener for vetoable changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be removed.
        """
        self.__component.removeVetoableChangeListener(name, listener)

    def set_property_value(self, name: str, value: Any) -> None:
        """
        Sets the value of a property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.
        """
        self.__component.setPropertyValue(name, value)

    # endregion XPropertySet


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.add_import(
        name="ooodev.adapter.beans.property_set_partial.PropertySetPartial",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=False,
        init_kind=2,
    )
    return builder
