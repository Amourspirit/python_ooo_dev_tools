from __future__ import annotations
from typing import Any
import uno

from com.sun.star.text import XWordCursor


from ooodev.adapter.text.word_cursor_partial import WordCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from .write_text_cursor import WriteTextCursor


class WriteWordCursor(WriteTextCursor, WordCursorPartial, StylePartial):
    """Represents a writer word cursor."""

    def __init__(self, owner: Any, component: XWordCursor) -> None:
        """
        Constructor

        Args:
            owner (Any): Owner of this component.
            component (XWordCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
        """
        WriteTextCursor.__init__(self, owner=owner, component=component)
        WordCursorPartial.__init__(self, component)  # type: ignore
        StylePartial.__init__(self, component=component)
