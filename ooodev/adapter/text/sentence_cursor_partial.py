from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.text import XSentenceCursor
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class SentenceCursorPartial:
    """
    Class for managing SentenceCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSentenceCursor) -> None:
        """
        Constructor

        Args:
            component (XSentenceCursor): UNO Component that implements ``com.sun.star.text.XSentenceCursor`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XSentenceCursor):
            raise mEx.MissingInterfaceError("XSentenceCursor")
        self.__component = component

    # region XSentenceCursor
    def is_start_of_sentence(self) -> bool:
        """
        Returns true if the cursor is at the start of a sentence.

        Returns:
            bool: ``True`` if the cursor is at the start of a sentence.
        """

        return self.__component.isStartOfSentence()

    def is_end_of_sentence(self) -> bool:
        """
        Returns true if the cursor is at the end of a sentence.

        Returns:
            bool: ``True`` if the cursor is at the end of a sentence.
        """

        return self.__component.isEndOfSentence()

    def goto_next_sentence(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the next sentence.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the next sentence. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoNextSentence(expand)

    def goto_previous_sentence(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the previous sentence.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the previous sentence. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoPreviousSentence(expand)

    def goto_start_of_sentence(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the start of the sentence.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the start of the sentence. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the start of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoStartOfSentence(expand)

    def goto_end_of_sentence(self, expand: bool = False) -> bool:
        """
        Moves the cursor to the end of the sentence.

        Args:
            expand (bool, optional): If ``True`` the range of the cursor will be expanded to the end of the sentence. Default is ``False``.

        Returns:
            bool: ``True`` if the cursor is now at the end of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """

        return self.__component.gotoEndOfSentence(expand)

    # endregion XSentenceCursor
