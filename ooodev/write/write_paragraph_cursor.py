from __future__ import annotations
from typing import Any
import uno

from com.sun.star.text import XParagraphCursor

from ooodev.adapter.text.paragraph_cursor_partial import ParagraphCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from .write_text_cursor import WriteTextCursor


class WriteParagraphCursor(WriteTextCursor, ParagraphCursorPartial, StylePartial):
    """Represents a writer Paragraph cursor."""

    def __init__(self, owner: Any, component: XParagraphCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Doc that owns this component.
            component (XParagraphCursor): A UNO object that supports ``com.sun.star.text.XParagraphCursor`` interface.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        WriteTextCursor.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        ParagraphCursorPartial.__init__(self, component, None)  # type: ignore
        StylePartial.__init__(self, component=component)
