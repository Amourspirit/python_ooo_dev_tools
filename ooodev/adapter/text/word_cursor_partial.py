from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.text import XWordCursor
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class WordCursorPartial:
    """
    Partial class for XWordCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XWordCursor, interface: UnoInterface | None = XWordCursor) -> None:
        """
        Constructor

        Args:
            component (XWordCursor): UNO Component that implements ``com.sun.star.text.XWordCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XWordCursor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XWordCursor

    def is_start_of_word(self) -> bool:
        """Returns True if the cursor is at the start of a word."""
        return self.__component.isStartOfWord()

    def is_end_of_word(self) -> bool:
        """Returns True if the cursor is at the end of a word."""
        return self.__component.isEndOfWord()

    def goto_next_word(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the next word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the end of the word. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor is moved, ``False`` otherwise.

        Note:
            The function returning ``True`` does not necessarily mean that the cursor is located at the next word,
            or any word at all! This may happen for example if it travels over empty paragraphs.
        """
        return self.__component.gotoNextWord(expand)

    def goto_previous_word(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the previous word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the start of the word. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor is moved, ``False`` otherwise.

        Note:
            The function returning ``True`` does not necessarily mean that the cursor is located at the previous word,
            or any word at all! This may happen for example if it travels over empty paragraphs.
        """
        return self.__component.gotoPreviousWord(expand)

    def goto_end_of_word(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the end of the current word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the end of the word. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor is moved, ``False`` otherwise.
        """
        return self.__component.gotoEndOfWord(expand)

    def goto_start_of_word(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the start of the current word.

        Args:
            expand (bool, optional): If ``True``, the selection is expanded to the start of the word. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor is moved, ``False`` otherwise.
        """
        return self.__component.gotoStartOfWord(expand)

    # endregion XWordCursor
