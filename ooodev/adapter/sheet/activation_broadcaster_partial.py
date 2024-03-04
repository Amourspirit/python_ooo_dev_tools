from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XActivationBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
