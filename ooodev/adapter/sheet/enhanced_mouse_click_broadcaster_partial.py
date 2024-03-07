from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XEnhancedMouseClickBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XEnhancedMouseClickHandler
    from ooodev.utils.type_var import UnoInterface


class EnhancedMouseClickBroadcasterPartial:
    """
    Partial Class for XEnhancedMouseClickBroadcaster.

    .. versionadded:: 0.20.0
    """

    def __init__(
        self,
        component: XEnhancedMouseClickBroadcaster,
        interface: UnoInterface | None = XEnhancedMouseClickBroadcaster,
    ) -> None:
        """
        Constructor

        Args:
            component (XEnhancedMouseClickBroadcaster): UNO Component that implements ``com.sun.star.sheet.XEnhancedMouseClickBroadcaster``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEnhancedMouseClickBroadcaster``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XEnhancedMouseClickBroadcaster
    def add_enhanced_mouse_click_handler(self, handler: XEnhancedMouseClickHandler) -> None:
        """
        Adds the specified mouse click handler to receive mouse click events from this source.

        Args:
            handler (XEnhancedMouseClickHandler): The mouse click handler to add.
        """
        self.__component.addEnhancedMouseClickHandler(handler)

    def remove_enhanced_mouse_click_handler(self, handler: XEnhancedMouseClickHandler) -> None:
        """
        Removes the specified mouse click handler so it does not receive mouse click events from this source anymore.

        Args:
            handler (XEnhancedMouseClickHandler): The mouse click handler to remove.
        """
        self.__component.removeEnhancedMouseClickHandler(handler)

    # endregion XEnhancedMouseClickBroadcaster
