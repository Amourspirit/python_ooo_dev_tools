from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.view import XScreenCursor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
