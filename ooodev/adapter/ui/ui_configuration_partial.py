from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.ui import XUIConfiguration

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.ui import XUIConfigurationListener
    from ooodev.utils.type_var import UnoInterface


class UIConfigurationPartial:
    """
    Partial Class for XUIConfiguration.
    """

    def __init__(self, component: XUIConfiguration, interface: UnoInterface | None = XUIConfiguration) -> None:
        """
        Constructor

        Args:
            component (XUIConfiguration): UNO Component that implements ``com.sun.star.ui.XUIConfiguration``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIConfiguration``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUIConfiguration
    def add_configuration_listener(self, listener: XUIConfigurationListener) -> None:
        """
        adds the specified listener to receive events when elements are changed, inserted or removed.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        self.__component.addConfigurationListener(listener)

    def remove_configuration_listener(self, listener: XUIConfigurationListener) -> None:
        """
        removes the specified listener so it does not receive any events from this user interface configuration manager.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        self.__component.removeConfigurationListener(listener)

    # endregion XUIConfiguration


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.ui.XUIConfiguration")
    return builder
