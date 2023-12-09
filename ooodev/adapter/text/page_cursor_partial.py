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

    def jump_to_first_page(self) -> bool:
        """
        Jumps to the first page.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToFirstPage()

    def jump_to_last_page(self) -> None:
        """Jumps to the last page."""
        self.__component.jumpToLastPage()

    def jump_to_next_page(self) -> bool:
        """
        Jumps to the next page.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToNextPage()

    def jump_to_previous_page(self) -> bool:
        """
        Jumps to the previous page.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToPreviousPage()

    def jump_to_start_of_page(self) -> bool:
        """
        Jumps to the start of the current page.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToStartOfPage()

    def jump_to_end_of_page(self) -> bool:
        """
        Jumps to the end of the current page.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToEndOfPage()

    def jump_to_page_(self, number: int) -> bool:
        """
        Jumps to the page with the given number.

        Args:
            number (int): Page number to jump to.

        Returns:
            bool: ``True`` if the jump was successful, ``False`` otherwise.
        """
        return self.__component.jumpToPage(number)

    # endregion XPageCursor
