from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XDataTransferProviderAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.datatransfer.clipboard import XClipboard
    from com.sun.star.datatransfer.dnd import XDragGestureRecognizer
    from com.sun.star.datatransfer.dnd import XDragSource
    from com.sun.star.datatransfer.dnd import XDropTarget
    from com.sun.star.awt import XWindow
    from ooodev.utils.type_var import UnoInterface


class DataTransferProviderAccessPartial:
    """
    Partial class for XDataTransferProviderAccess.
    """

    def __init__(
        self, component: XDataTransferProviderAccess, interface: UnoInterface | None = XDataTransferProviderAccess
    ) -> None:
        """
        Constructor

        Args:
            component (XDataTransferProviderAccess): UNO Component that implements ``com.sun.star.awt.XDataTransferProviderAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataTransferProviderAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDataTransferProviderAccess
    def get_clipboard(self, clipboard_name: str) -> XClipboard:
        """
        Gets the specified clipboard.
        """
        return self.__component.getClipboard(clipboard_name)

    def get_drag_gesture_recognizer(self, window: XWindow) -> XDragGestureRecognizer:
        """
        Gets the drag gesture recognizer of the specified window.
        """
        return self.__component.getDragGestureRecognizer(window)

    def get_drag_source(self, window: XWindow) -> XDragSource:
        """
        Gets the drag source of the specified window.
        """
        return self.__component.getDragSource(window)

    def get_drop_target(self, window: XWindow) -> XDropTarget:
        """
        Gets the drop target of the specified window.
        """
        return self.__component.getDropTarget(window)

    # endregion XDataTransferProviderAccess
