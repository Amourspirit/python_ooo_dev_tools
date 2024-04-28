from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XFastPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FastPropertySetPartial:
    """
    Partial class for XFastPropertySet.
    """

    def __init__(self, component: XFastPropertySet, interface: UnoInterface | None = XFastPropertySet) -> None:
        """
        Constructor

        Args:
            component (XFastPropertySet): UNO Component that implements ``com.sun.star.beans.XFastPropertySet`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFastPropertySet``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFastPropertySet
    def get_fast_property_value(self, handle: int) -> Any:
        """
        returns the value of the property with the specified name.

        Args:
            handle (int): The implementation handle of the implementation for the property.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.getFastPropertyValue(handle)

    def set_fast_property_value(self, handle: int, value: Any) -> None:
        """
        Sets the value to the property with the specified name.

        Args:
            handle (int): The implementation handle of the implementation for the property.
            value (Any): The new value for the property.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.setFastPropertyValue(handle, value)

    # endregion XFastPropertySet


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
    builder.auto_add_interface("com.sun.star.beans import XFastPropertySet")
    return builder
