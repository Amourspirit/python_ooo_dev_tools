from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XHierarchicalPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XHierarchicalPropertySetInfo


class HierarchicalPropertySetPartial:
    """
    Partial class for XHierarchicalPropertySet.
    """

    def __init__(
        self, component: XHierarchicalPropertySet, interface: UnoInterface | None = XHierarchicalPropertySet
    ) -> None:
        """
        Constructor

        Args:
            component (XHierarchicalPropertySet): UNO Component that implements ``com.sun.star.bean.XHierarchicalPropertySet`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XHierarchicalPropertySet``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XHierarchicalPropertySet

    def get_hierarchical_property_set_info(self) -> XHierarchicalPropertySetInfo:
        """
        Gets information about the hierarchy of properties.
        """
        return self.__component.getHierarchicalPropertySetInfo()

    def get_hierarchical_property_value(self, name: str) -> Any:
        """
        Gets the value of the property with the specified name.

        Args:
            name (str): The Hierarchical Property Name of the property.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.getHierarchicalPropertyValue(name)

    def set_hierarchical_property_value(self, name: str, value: Any) -> None:
        """
        Sets the value of the property with the specified nested name.

        Args:
            name (str): The Hierarchical Property Name of the property.
            value (Any): The value to be set.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.setHierarchicalPropertyValue(name, value)

    # endregion XHierarchicalPropertySet


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
    builder.auto_add_interface("com.sun.star.beans.XHierarchicalPropertySet", False)
    return builder
