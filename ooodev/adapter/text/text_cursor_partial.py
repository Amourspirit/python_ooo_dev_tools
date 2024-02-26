from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.text import XTextCursor

from ooodev.adapter.text import text_range_partial as mTextRangeComp

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from ooodev.utils.type_var import UnoInterface


class TextCursorPartial(mTextRangeComp.TextRangePartial):
    """
    Partial class for XTextCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextCursor, interface: UnoInterface | None = XTextCursor) -> None:
        """
        Constructor

        Args:
            component (XTextCursor): UNO Component that implements ``com.sun.star.text.XTextCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextCursor``.
        """
        mTextRangeComp.TextRangePartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XTextCursor
    def collapse_to_end(self) -> None:
        """Sets the end of the position to the start."""
        self.__component.collapseToEnd()

    def collapse_to_start(self) -> None:
        """Sets the start of the position to the end."""
        self.__component.collapseToStart()

    def is_collapsed(self) -> bool:
        """Returns ``True`` if the cursor is collapsed."""
        return self.__component.isCollapsed()

    def go_left(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor left by the given number of units.

        Args:
            count (int): Number of units to move.
            expand (bool, optional): ``True`` to expand the selection. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved left, ``False`` otherwise.

        Note:
            Even if the command was not completed successfully it may be completed partially.
            E.g. if it was required to move 5 characters but it is only possible to move 3 FALSE will be returned and the cursor moves only those 3 characters.
        """
        return self.__component.goLeft(count, expand)

    def go_right(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor right by the given number of units.

        Args:
            count (int): Number of units to move.
            expand (bool, optional): ``True`` to expand the selection. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved right, ``False`` otherwise.

        Note:
            Even if the command was not completed successfully it may be completed partially.
            E.g. if it was required to move 5 characters but it is only possible to move 3 FALSE will be returned and the cursor moves only those 3 characters.
        """
        return self.__component.goRight(count, expand)

    def goto_end(self, expand: bool = False) -> None:
        """
        Moves the cursor to the end of the document.

        Args:
            expand (bool, optional): ``True`` to expand the selection. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoEnd(expand)

    def goto_range(self, range: XTextRange, expand: bool = False) -> None:
        """
        Moves or expands the cursor to a specified TextRange.

        Args:
            range (XTextRange): Range to move to.
            expand (bool, optional): ``True`` to expand the selection. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoRange(range, expand)

    def goto_start(self, expand: bool = False) -> None:
        """
        Moves the cursor to the start of the document.

        Args:
            expand (bool, optional): ``True`` to expand the selection. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoStart(expand)

    # endregion XTextCursor
