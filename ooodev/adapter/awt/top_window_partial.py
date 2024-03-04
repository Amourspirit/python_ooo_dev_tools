from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XTopWindow

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XTopWindowListener
    from com.sun.star.awt import XMenuBar
    from ooodev.utils.type_var import UnoInterface


class TopWindowPartial:
    """
    Partial class for XTopWindow.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTopWindow, interface: UnoInterface | None = XTopWindow) -> None:
        """
        Constructor

        Args:
            component (XTopWindow): UNO Component that implements ``com.sun.star.awt.XTopWindow`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTopWindow``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTopWindow
    def add_top_window_listener(self, listener: XTopWindowListener) -> None:
        """
        Adds the specified top window listener to receive window events from this window.
        """
        self.__component.addTopWindowListener(listener)

    def remove_top_window_listener(self, listener: XTopWindowListener) -> None:
        """
        Removes the specified top window listener so that it no longer receives window events from this window.
        """
        self.__component.removeTopWindowListener(listener)

    def set_menu_bar(self, menu: XMenuBar) -> None:
        """
        Sets a menu bar.
        """
        self.__component.setMenuBar(menu)

    def to_back(self) -> None:
        """
        Places this window at the bottom of the stacking order and makes the corresponding adjustment to other visible windows.
        """
        self.__component.toBack()

    def to_front(self) -> None:
        """
        Places this window at the top of the stacking order and shows it in front of any other windows.
        """
        self.__component.toFront()

    # endregion XTopWindow
