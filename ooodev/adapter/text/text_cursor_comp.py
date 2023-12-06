from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from .text_range_comp import TextRangeComp


if TYPE_CHECKING:
    from com.sun.star.text import TextCursor  # service
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextRange


class TextCursorComp(TextRangeComp):
    """
    Class for managing TextCursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextCursor) -> None:
        """
        Constructor

        Args:
            component (TextCursor): UNO TextCursor Component that supports ``com.sun.star.text.TextCursor`` service.
        """

        TextCursorComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextCursor",)

    # endregion Overrides

    # region Methods
    # region XTextCursor
    def collapse_to_end(self) -> None:
        """Sets the end of the position to the start."""
        self.component.collapseToEnd()

    def collapse_to_start(self) -> None:
        """Sets the start of the position to the end."""
        self.component.collapseToStart()

    def is_collapsed(self) -> bool:
        """Returns True if the cursor is collapsed."""
        return self.component.isCollapsed()

    def go_left(self, count: int, expand: bool) -> None:
        """Moves the cursor left by the given number of units."""
        self.component.goLeft(count, expand)

    def go_right(self, count: int, expand: bool) -> None:
        """Moves the cursor right by the given number of units."""
        self.component.goRight(count, expand)

    def goto_end(self, expand: bool) -> None:
        """Moves the cursor to the end of the document."""
        self.component.gotoEnd(expand)

    def goto_range(self, range: XTextRange, expand: bool) -> None:
        """Moves the cursor to the given range."""
        self.component.gotoRange(range, expand)

    def goto_start(self, expand: bool) -> None:
        """Moves the cursor to the start of the document."""
        self.component.gotoStart(expand)

    # endregion XTextCursor

    # region XParagraphCursor

    def goto_end_of_paragraph(self, expand: bool) -> bool:
        """
        Moves the cursor to the end of the current paragraph.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the end of the paragraph.

        Returns:
            bool: ``True`` if the cursor is now at the end of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoEndOfParagraph(expand)

    def goto_start_of_paragraph(self, expand: bool) -> bool:
        """
        Moves the cursor to the start of the current paragraph.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the start of the paragraph.

        Returns:
            bool: ``True`` if the cursor is now at the start of a paragraph, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoStartOfParagraph(expand)

    def goto_next_paragraph(self, expand: bool) -> bool:
        """
        Moves the cursor to the next paragraph.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the next paragraph.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.
        """
        return self.component.gotoNextParagraph(expand)

    def goto_previous_paragraph(self, expand: bool) -> bool:
        """
        Moves the cursor to the previous paragraph.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the previous paragraph.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.
        """
        return self.component.gotoPreviousParagraph(expand)

    def is_end_of_paragraph(self) -> bool:
        """Returns ``True`` if the cursor is at the end of a paragraph."""
        return self.component.isEndOfParagraph()

    def is_start_of_paragraph(self) -> bool:
        """Returns ``True`` if the cursor is at the start of a paragraph."""
        return self.component.isStartOfParagraph()

    # endregion XParagraphCursor

    # region XWordCursor
    def is_start_of_word(self) -> bool:
        """Returns ``True`` if the cursor is at the start of a word."""
        return self.component.isStartOfWord()

    def is_end_of_word(self) -> bool:
        """Returns ``True`` if the cursor is at the end of a word."""
        return self.component.isEndOfWord()

    def goto_next_word(self, expand: bool) -> bool:
        """
        Moves the cursor to the next word.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the next word.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.

        Note:
            When ``True`` is returned it does not necessarily mean that the cursor is located at the next word,
            or any word at all! This may happen for example if it travels over empty paragraphs.
        """
        return self.component.gotoNextWord(expand)

    def goto_previous_word(self, expand: bool) -> bool:
        """
        Moves the cursor to the previous word.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the previous word.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.

        Note:
            When ``True`` is returned it does not necessarily mean that the cursor is located at the previous word,
            or any word at all! This may happen for example if it travels over empty paragraphs.
        """
        return self.component.gotoPreviousWord(expand)

    def goto_start_of_word(self, expand: bool) -> bool:
        """
        Moves the cursor to the start of the current word.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the start of the word.

        Returns:
            bool: ``True`` if the cursor is now at the start of a word, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoStartOfWord(expand)

    def goto_end_of_word(self, expand: bool) -> bool:
        """
        Moves the cursor to the end of the current word.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the end of the word.

        Returns:
            bool: ``True`` if the cursor is now at the end of a word, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoEndOfWord(expand)

    # endregion XWordCursor

    # region XSentenceCursor
    def is_start_of_sentence(self) -> bool:
        """Returns ``True`` if the cursor is at the start of a sentence."""
        return self.component.isStartOfSentence()

    def is_end_of_sentence(self) -> bool:
        """Returns ``True`` if the cursor is at the end of a sentence."""
        return self.component.isEndOfSentence()

    def goto_next_sentence(self, expand: bool) -> bool:
        """
        Moves the cursor to the next sentence.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the next sentence.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.
        """
        return self.component.gotoNextSentence(expand)

    def goto_previous_sentence(self, expand: bool) -> bool:
        """
        Moves the cursor to the previous sentence.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the previous sentence.

        Returns:
            bool: ``True`` if the cursor was moved. It returns ``False`` it the cursor can not advance further.
        """
        return self.component.gotoPreviousSentence(expand)

    def goto_start_of_sentence(self, expand: bool) -> bool:
        """
        Moves the cursor to the start of the current sentence.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the start of the sentence.

        Returns:
            bool: ``True`` if the cursor is now at the start of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoStartOfSentence(expand)

    def goto_end_of_sentence(self, expand: bool) -> bool:
        """
        Moves the cursor to the end of the current sentence.

        Args:
            expand (bool): If ``True`` the range of the cursor will be expanded to the end of the sentence.

        Returns:
            bool: ``True`` if the cursor is now at the end of a sentence, ``False`` otherwise. If ``False`` is returned the cursor will remain at its original position.
        """
        return self.component.gotoEndOfSentence(expand)

    # endregion XSentenceCursor

    # endregion Methods

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextCursor:
            """Sheet Cell Cursor Component"""
            return cast("TextCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
