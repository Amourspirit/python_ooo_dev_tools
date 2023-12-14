from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.view import XScreenCursor

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ScreenCursorPartial:
    """
    Partial class for XSentenceCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XScreenCursor, interface: UnoInterface | None = XScreenCursor) -> None:
        """
        Constructor

        Args:
            component (XScreenCursor): UNO Component that implements ``com.sun.star.view.XScreenCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XScreenCursor``.
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

    # region XScreenCursor
    def screen_down(self) -> None:
        """
        Scrolls the view forward by one visible page.
        """
        self.__component.screenDown()

    def screen_up(self) -> None:
        """
        Scrolls the view backward by one visible page.
        """
        self.__component.screenUp()

    # endregion XScreenCursor
