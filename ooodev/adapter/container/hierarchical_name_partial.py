from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XHierarchicalName

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface


class HierarchicalNamePartial:
    """
    Partial class for XHierarchicalName.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XHierarchicalName, interface: UnoInterface | None = XHierarchicalName) -> None:
        """
        Constructor

        Args:
            component (XHierarchicalName): UNO Component that implements ``com.sun.star.container.XHierarchicalName`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XHierarchicalName``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XHierarchicalName
    def compose_hierarchical_name(self, relative_name: str) -> str:
        """
        builds the hierarchical name of an object, given a relative name

        Can be used to find the name of a descendant object in the hierarchy without actually accessing it.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.NoSupportException: ``NoSupportException``
        """
        return self.__component.composeHierarchicalName(relative_name)

    def get_hierarchical_name(self) -> str:
        """
        returns the hierarchical name of the object within its parent container
        """
        return self.__component.getHierarchicalName()

    # endregion XHierarchicalName


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
        name="ooodev.adapter.container.hierarchical_name_partial.HierarchicalNamePartial",
        uno_name="com.sun.star.container.XHierarchicalName",
        optional=False,
        init_kind=2,
    )
    return builder
