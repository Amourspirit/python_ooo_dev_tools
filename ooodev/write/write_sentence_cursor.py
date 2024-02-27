from __future__ import annotations
from typing import Any
import uno

from com.sun.star.text import XSentenceCursor

from ooodev.adapter.text.sentence_cursor_partial import SentenceCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.write_text_cursor import WriteTextCursor


class WriteSentenceCursor(WriteTextCursor, SentenceCursorPartial, StylePartial):
    """Represents a writer Sentence cursor."""

    def __init__(self, owner: Any, component: XSentenceCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (Any): Doc that owns this component.
            component (XSentenceCursor): A UNO object that supports ``com.sun.star.text.XSentenceCursor`` interface.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WriteTextCursor.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        SentenceCursorPartial.__init__(self, component, None)  # type: ignore
        StylePartial.__init__(self, component=component)
