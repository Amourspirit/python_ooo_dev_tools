from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextRange

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from . import text_range_partial as mTextRangeComp


class TextCursorPartial(mTextRangeComp.TextRangePartial):
    """
    Class for managing TextCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextCursor) -> None:
        """
        Constructor

        Args:
            component (XTextCursor): UNO Component that implements ``com.sun.star.text.XTextCursor`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XTextCursor):
            raise mEx.MissingInterfaceError("XTextCursor")
        self.__component = component

    # region XTextCursor
    def collapse_to_end(self) -> None:
        """Sets the end of the position to the start."""
        self.__component.collapseToEnd()

    def collapse_to_start(self) -> None:
        """Sets the start of the position to the end."""
        self.__component.collapseToStart()

    def is_collapsed(self) -> bool:
        """Returns True if the cursor is collapsed."""
        return self.__component.isCollapsed()

    def go_left(self, count: int, expand: bool) -> None:
        """Moves the cursor left by the given number of units."""
        self.__component.goLeft(count, expand)

    def go_right(self, count: int, expand: bool) -> None:
        """Moves the cursor right by the given number of units."""
        self.__component.goRight(count, expand)

    def goto_end(self, expand: bool) -> None:
        """Moves the cursor to the end of the document."""
        self.__component.gotoEnd(expand)

    def goto_range(self, range: XTextRange, expand: bool) -> None:
        """Moves the cursor to the given range."""
        self.__component.gotoRange(range, expand)

    def goto_start(self, expand: bool) -> None:
        """Moves the cursor to the start of the document."""
        self.__component.gotoStart(expand)

    # endregion XTextCursor
