from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XExtendedToolkit

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XFocusListener
    from com.sun.star.awt import XKeyHandler
    from com.sun.star.awt import XTopWindowListener
    from com.sun.star.uno import XInterface
    from com.sun.star.awt import XTopWindow
    from ooodev.utils.type_var import UnoInterface


class ExtendedToolkitPartial:
    """
    Partial class for XExtendedToolkit.
    """

    def __init__(self, component: XExtendedToolkit, interface: UnoInterface | None = XExtendedToolkit) -> None:
        """
        Constructor

        Args:
            component (XExtendedToolkit): UNO Component that implements ``com.sun.star.awt.XExtendedToolkit`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XExtendedToolkit``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XExtendedToolkit
    def add_focus_listener(self, listener: XFocusListener) -> None:
        """
        Add a new listener that is called on com.sun.star.awt.FocusEvent.

        Use this focus broadcaster to keep track of the object that currently has the input focus.
        """
        self.__component.addFocusListener(listener)

    def add_key_handler(self, handler: XKeyHandler) -> None:
        """
        Add a new listener that is called on com.sun.star.awt.KeyEvent.

        Every listener is given the opportunity to consume the event, i.e. prevent the not yet called listeners from being called.
        """
        self.__component.addKeyHandler(handler)

    def add_top_window_listener(self, listener: XTopWindowListener) -> None:
        """
        Add a new listener that is called for events that involve com.sun.star.awt.XTopWindow.

        After having obtained the current list of existing top-level windows you can keep this list up-to-date by listening to opened or closed top-level windows. Wait for activations or deactivations of top-level windows to keep track of the currently active frame.
        """
        self.__component.addTopWindowListener(listener)

    def fire_focus_gained(self, source: XInterface) -> None:
        """
        Broadcasts the a focusGained on all registered focus listeners.
        """
        self.__component.fireFocusGained(source)

    def fire_focus_lost(self, source: XInterface) -> None:
        """
        Broadcasts the a focusGained on all registered focus listeners.
        """
        self.__component.fireFocusLost(source)

    def get_active_top_window(self) -> XTopWindow:
        """
        Return the currently active top-level window, i.e.

        which has currently the input focus.
        """
        return self.__component.getActiveTopWindow()

    def get_top_window(self, idx: int) -> XTopWindow:
        """
        Return a reference to the specified top-level window.

        Note that the number of top-level windows may change between a call to getTopWindowCount() and successive calls to this function.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getTopWindow(idx)

    def get_top_window_count(self) -> int:
        """
        This function returns the number of currently existing top-level windows.
        """
        return self.__component.getTopWindowCount()

    def remove_focus_listener(self, listener: XFocusListener) -> None:
        """
        Remove the specified listener from the list of listeners.
        """
        self.__component.removeFocusListener(listener)

    def remove_key_handler(self, handler: XKeyHandler) -> None:
        """
        Remove the specified listener from the list of listeners.
        """
        self.__component.removeKeyHandler(handler)

    def remove_top_window_listener(self, xListener: XTopWindowListener) -> None:
        """
        Remove the specified listener from the list of listeners.
        """
        self.__component.removeTopWindowListener(xListener)

    # endregion XExtendedToolkit
