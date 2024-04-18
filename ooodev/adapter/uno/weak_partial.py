from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.uno import XWeak

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.uno import XAdapter
    from ooodev.utils.type_var import UnoInterface


class WeakPartial:
    """
    Partial Class for XUIConfiguration.
    """

    def __init__(self, component: XWeak, interface: UnoInterface | None = XWeak) -> None:
        """
        Constructor

        Args:
            component (XWeak): UNO Component that implements ``com.sun.star.uno.XWeak``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XWeak``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XWeak
    def query_adapter(self) -> XAdapter:
        """
        Queries the weak adapter.

        It is important that the adapter must know, but not hold the adapted object. If the adapted object dies, all references to the adapter have to be notified to release the adapter.
        """
        return self.__component.queryAdapter()

    # endregion XWeak


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.uno.XWeak")
    return builder
