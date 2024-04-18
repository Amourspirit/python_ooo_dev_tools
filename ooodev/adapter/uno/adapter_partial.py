from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.uno import XAdapter

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.uno import XReference
    from ooodev.utils.type_var import UnoInterface


class AdapterPartial:
    """
    Partial Class for XAdapter.
    """

    def __init__(self, component: XAdapter, interface: UnoInterface | None = XAdapter) -> None:
        """
        Constructor

        Args:
            component (XAdapter): UNO Component that implements ``com.sun.star.uno.XAdapter``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAdapter``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XAdapter
    def add_reference(self, ref: XReference) -> None:
        """
        adds a reference to the adapter.

        All added references are called when the adapted object dies.
        """
        self.__component.addReference(ref)

    def query_adapted(self) -> XInterface:
        """
        queries the adapted object if it is alive.
        """
        return self.__component.queryAdapted()

    def remove_reference(self, ref: XReference) -> None:
        """
        removes a reference from the adapter.
        """
        self.__component.removeReference(ref)

    # endregion XAdapter


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.uno.XAdapter")
    return builder
