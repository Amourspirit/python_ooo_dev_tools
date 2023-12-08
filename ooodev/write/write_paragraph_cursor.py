from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XParagraphCursor


from ooodev.adapter.text.paragraph_cursor_partial import ParagraphCursorPartial

from .write_text_cursor import WriteTextCursor


class WriteParagraphCursor(WriteTextCursor, ParagraphCursorPartial):
    """Represents a writer Paragraph cursor."""

    def __init__(self, owner: Any, component: XParagraphCursor) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Doc that owns this component.
            component (XParagraphCursor): A UNO object that supports ``com.sun.star.text.XParagraphCursor`` interface.
        """
        WriteTextCursor.__init__(self, owner=owner, component=component)
        ParagraphCursorPartial.__init__(self, component)  # type: ignore
