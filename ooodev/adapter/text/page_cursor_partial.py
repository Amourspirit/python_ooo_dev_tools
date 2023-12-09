from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.text import XPageCursor
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PageCursorPartial:
    """
    Partial class for XPageCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPageCursor, interface: UnoInterface | None = XPageCursor) -> None:
        """
        Constructor

        Args:
            component (XPageCursor): UNO Component that implements ``com.sun.star.text.XPageCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPageCursor``.
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
