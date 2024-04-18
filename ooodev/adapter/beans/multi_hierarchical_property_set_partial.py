from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XMultiHierarchicalPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XHierarchicalPropertySetInfo


class MultiHierarchicalPropertySetPartial:
    """
    Partial class for XMultiHierarchicalPropertySet.
    """

    def __init__(
        self, component: XMultiHierarchicalPropertySet, interface: UnoInterface | None = XMultiHierarchicalPropertySet
    ) -> None:
        """
        Constructor

        Args:
            component (XMultiHierarchicalPropertySet): UNO Component that implements ``com.sun.star.bean.XMultiHierarchicalPropertySet`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiHierarchicalPropertySet``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMultiHierarchicalPropertySet

    def get_hierarchical_property_set_info(self) -> XHierarchicalPropertySetInfo:
        """
        Gets information about the hierarchy of properties
        """
        return self.__component.getHierarchicalPropertySetInfo()

    def get_hierarchical_property_values(self, *names: str) -> Tuple[Any, ...]:
        """
        The order of the values in the returned sequence will be the same as the order of the names in the argument.

        Unknown properties are ignored, in their place NULL will be returned.

        Args:
            names (str): The names of the properties.

        Raises:
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.getHierarchicalPropertyValues(names)

    def set_hierarchical_property_values(self, names: Tuple[str, ...], values: Tuple[Any, ...]) -> None:
        """
        sets the values of the properties with the specified nested names.

        The values of the properties must change before bound events are fired. The values of constrained properties should change after the vetoable events are fired, if no exception occurs.
        Unknown properties are ignored.

        Similar to ``set_hierarchical_prop_values()`` but with tuple arguments.

        Args:
            names (Tuple[str, ...]): The hierarchical property names of the properties.
            values (Tuple[Any, ...]): The values to be set.

        Raises:
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.setHierarchicalPropertyValues(names, values)

    # endregion XMultiHierarchicalPropertySet

    def set_hierarchical_prop_values(self, **kwargs) -> None:
        """
        Sets the values of the properties with the specified nested names.

        The values of the properties must change before bound events are fired. The values of constrained properties should change after the vetoable events are fired, if no exception occurs.
        Unknown properties are ignored.

        Similar to ``set_hierarchical_property_values()`` but with key, value arguments.

        Args:
            **kwargs: The hierarchical property names and values of the properties.

        Raises:
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        names, values = zip(*kwargs.items())
        self.__component.setHierarchicalPropertyValues(names, values)


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
