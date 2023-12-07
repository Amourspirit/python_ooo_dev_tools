from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextViewCursor
    from com.sun.star.awt import Point

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from . import text_cursor_partial as mTextCursorPartial


class TextViewCursorPartial(mTextCursorPartial.TextCursorPartial):
    """
    Class for TextViewCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextViewCursor) -> None:
        """
        Constructor

        Args:
            component (XTextViewCursor): UNO Component that implements ``com.sun.star.text.XTextViewCursor`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XTextViewCursor):
            raise mEx.MissingInterfaceError("XTextViewCursor")
        self.__component = component

    # region XTextViewCursor
    def is_visible(self) -> bool:
        """Returns True if the cursor is visible."""
        return self.__component.isVisible()

    def set_visible(self, visible: bool) -> None:
        """Sets the visibility of the cursor."""
        self.__component.setVisible(visible)

    def get_position(self) -> Point:
        """Returns the position of the cursor."""
        return self.__component.getPosition()

    # endregion XTextViewCursor
