from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XSentenceCursor

from ooodev.adapter.text.sentence_cursor_partial import SentenceCursorPartial
from ooodev.format.inner.style_partial import StylePartial

from .write_text_cursor import WriteTextCursor


class WriteSentenceCursor(WriteTextCursor, SentenceCursorPartial, StylePartial):
    """Represents a writer Sentence cursor."""

    def __init__(self, owner: Any, component: XSentenceCursor) -> None:
        """
        Constructor

        Args:
            owner (Any): Doc that owns this component.
            component (XSentenceCursor): A UNO object that supports ``com.sun.star.text.XSentenceCursor`` interface.
        """
        WriteTextCursor.__init__(self, owner=owner, component=component)
        SentenceCursorPartial.__init__(self, component)  # type: ignore
        StylePartial.__init__(self, component=component)
