from __future__ import annotations
import uno

from com.sun.star.text import XWordCursor
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class WordCursorPartial:
    """
    Class for managing WordCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XWordCursor) -> None:
        """
        Constructor

        Args:
            component (XWordCursor): UNO Component that implements ``com.sun.star.text.XWordCursor`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XWordCursor):
            raise mEx.MissingInterfaceError("XWordCursor")
        self.__component = component

    # region XWordCursor

    def is_start_of_word(self) -> bool:
        """Returns True if the cursor is at the start of a word."""
        return self.__component.isStartOfWord()

    def is_end_of_word(self) -> bool:
        """Returns True if the cursor is at the end of a word."""
        return self.__component.isEndOfWord()

    def goto_next_word(self, expand: bool = False) -> None:
        """
        Moves the cursor to the next word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the end of the word. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoNextWord(expand)

    def goto_previous_word(self, expand: bool = False) -> None:
        """
        Moves the cursor to the previous word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the start of the word. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoPreviousWord(expand)

    def goto_end_of_word(self, expand: bool = False) -> None:
        """
        Moves the cursor to the end of the current word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the end of the word. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoEndOfWord(expand)

    def goto_start_of_word(self, expand: bool = False) -> None:
        """
        Moves the cursor to the start of the current word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the start of the word. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoStartOfWord(expand)

    # endregion XWordCursor
