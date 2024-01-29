from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XActivationBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import XActivationEventListener
    from ooodev.utils.type_var import UnoInterface


class ActivationBroadcasterPartial:
    """
    Partial Class for XActivationBroadcaster.

    .. versionadded:: 0.20.0
    """

    def __init__(
        self,
        component: XActivationBroadcaster,
        interface: UnoInterface | None = XActivationBroadcaster,
    ) -> None:
        """
        Constructor

        Args:
            component (XActivationBroadcaster): UNO Component that implements ``com.sun.star.sheet.XActivationBroadcaster``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XActivationBroadcaster``.
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

    # region XActivationBroadcaster
    def add_activation_event_listener(self, listener: XActivationEventListener) -> None:
        """
        Adds the specified activation event listener to receive activation events from this source.

        Args:
            listener (XActivationEventListener): The activation event listener to add.
        """
        self.__component.addActivationEventListener(listener)

    def remove_activation_event_listener(self, listener: XActivationEventListener) -> None:
        """
        Removes the specified activation event listener so that it no longer receives activation events from this source.

        Args:
            listener (XActivationEventListener): The activation event listener to remove.
        """
        self.__component.removeActivationEventListener(listener)

    # endregion XActivationBroadcaster
