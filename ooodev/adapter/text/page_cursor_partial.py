from __future__ import annotations
import uno

from com.sun.star.text import XPageCursor
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class PageCursorPartial:
    """
    Class for managing PageCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPageCursor) -> None:
        """
        Constructor

        Args:
            component (XPageCursor): UNO Component that implements ``com.sun.star.text.XPageCursor`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XPageCursor):
            raise mEx.MissingInterfaceError("XPageCursor")
        self.__component = component

    # region XPageCursor
    def get_page(self) -> int:
        """Returns the current page number."""
        return self.__component.getPage()

    def jump_to_first_page(self) -> None:
        """Jumps to the first page."""
        self.__component.jumpToFirstPage()

    def jump_to_last_page(self) -> None:
        """Jumps to the last page."""
        self.__component.jumpToLastPage()

    def jump_to_next_page(self) -> None:
        """Jumps to the next page."""
        self.__component.jumpToNextPage()

    def jump_to_previous_page(self) -> None:
        """Jumps to the previous page."""
        self.__component.jumpToPreviousPage()

    def jump_to_start_of_page(self) -> None:
        """Jumps to the start of the current page."""
        self.__component.jumpToStartOfPage()

    def jump_to_end_of_page(self) -> None:
        """Jumps to the end of the current page."""
        self.__component.jumpToEndOfPage()

    def jump_to_page_(self, number: int) -> None:
        """
        Jumps to the page with the given number.

        Args:
            number (int): Page number to jump to.
        """
        self.__component.jumpToPage(number)

    # endregion XPageCursor
