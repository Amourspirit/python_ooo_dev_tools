from __future__ import annotations
from typing import Any
import uno

from com.sun.star.text import XWordCursor


from ooodev.adapter.text.word_cursor_partial import WordCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from .write_text_cursor import WriteTextCursor


class WriteWordCursor(WriteTextCursor, WordCursorPartial, StylePartial):
    """Represents a writer word cursor."""

    def __init__(self, owner: Any, component: XWordCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (Any): Owner of this component.
            component (XWordCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        WriteTextCursor.__init__(self, owner=owner, component=component, lo_inst=self._lo_inst)
        WordCursorPartial.__init__(self, component, None)  # type: ignore
        StylePartial.__init__(self, component=component)
