from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XLayoutManagerEventBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.frame import XLayoutManagerListener
    from ooodev.utils.type_var import UnoInterface


class LayoutManagerEventBroadcasterPartial:
    """
    Partial class for XLayoutManagerEventBroadcaster.
    """

    def __init__(
        self,
        component: XLayoutManagerEventBroadcaster,
        interface: UnoInterface | None = XLayoutManagerEventBroadcaster,
    ) -> None:
        """
        Constructor

        Args:
            component (XLayoutManagerEventBroadcaster): UNO Component that implements ``com.sun.star.frame.XLayoutManagerEventBroadcaster`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLayoutManagerEventBroadcaster``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLayoutManagerEventBroadcaster
    def add_layout_manager_event_listener(self, listener: XLayoutManagerListener) -> None:
        """
        Adds a layout manager event listener to the object's listener list.
        """
        self.__component.addLayoutManagerEventListener(listener)

    def remove_layout_manager_event_listener(self, listener: XLayoutManagerListener) -> None:
        """
        Removes a layout manager event listener from the object's listener list.
        """
        self.__component.removeLayoutManagerEventListener(listener)

    # endregion XLayoutManagerEventBroadcaster
