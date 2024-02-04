from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.lang import XComponent

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.lang import XEventListener


class ComponentPartial:
    """
    Partial class for XComponent.
    """

    def __init__(self, component: XComponent, interface: UnoInterface | None = XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO Component that implements ``com.sun.star.container.XComponent`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XComponent``.
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

    # region XComponent
    def add_event_listener(self, listener: XEventListener) -> None:
        """
        Adds an event listener to the component.

        Args:
            listener (XEventListener): The event listener to be added.
        """
        self.__component.addEventListener(listener)

    def remove_event_listener(self, listener: XEventListener) -> None:
        """
        Removes an event listener from the component.

        Args:
            listener (XEventListener): The event listener to be removed.
        """
        self.__component.removeEventListener(listener)

    def dispose(self) -> None:
        """
        Disposes the component.
        """
        self.__component.dispose()

    # endregion XComponent
