from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XSentenceCursor

from ooodev.adapter.text.sentence_cursor_partial import SentenceCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from .write_text_cursor import WriteTextCursor


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
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        WriteTextCursor.__init__(self, owner=owner, component=component, lo_inst=self._lo_inst)
        SentenceCursorPartial.__init__(self, component, None)  # type: ignore
        StylePartial.__init__(self, component=component)
