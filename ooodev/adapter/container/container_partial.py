from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XContainer

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.container import XContainerListener


class ContainerPartial:
    """
    Partial class for XContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XContainer, interface: UnoInterface | None = XContainer) -> None:
        """
        Constructor

        Args:
            component (XContainer): UNO Component that implements ``com.sun.star.container.XContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContainer``.
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

    # region XContainer
    def add_container_listener(self, listener: XContainerListener) -> None:
        """
        Adds the specified listener to receive events when elements are inserted or removed.

        Args:
            listener (XContainerListener): The listener to be added.
        """
        self.__component.addContainerListener(listener)

    def remove_container_listener(self, listener: XContainerListener) -> None:
        """
        Removes the specified listener so it does not receive any events from this container.

        Args:
            listener (XContainerListener): The listener to be removed.
        """
        self.__component.removeContainerListener(listener)

    # endregion XContainer
