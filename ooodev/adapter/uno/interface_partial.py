from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.uno import XInterface

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.type_var import UnoInterface


class InterfacePartial:
    """
    Partial Class for XUIConfiguration.
    """

    def __init__(self, component: XInterface, interface: UnoInterface | None = XInterface) -> None:
        """
        Constructor

        Args:
            component (XInterface): UNO Component that implements ``com.sun.star.uno.XInterface``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XInterface``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XInterface
    def query_interface(self, aType: Any) -> Any:
        return self.__component.queryInterface(aType)

    def release(self) -> None:
        """
        decreases the reference counter by one.

        When the reference counter reaches 0, the object gets deleted.

        Calling release() on the object is often called releasing or clearing the reference to an object.
        """
        self.__component.release()

    # endregion XInterface


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.uno.XInterface")
    return builder
