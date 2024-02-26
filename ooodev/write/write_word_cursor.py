from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.adapter.text.word_cursor_partial import WordCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.write_text_cursor import WriteTextCursor

if TYPE_CHECKING:
    from com.sun.star.text import XWordCursor
    from ooodev.loader.inst.lo_inst import LoInst


class WriteWordCursor(WriteTextCursor, WriteDocPropPartial, WordCursorPartial, StylePartial):
    """Represents a writer word cursor."""

    def __init__(self, owner: Any, component: XWordCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (Any): Owner of this component.
            component (XWordCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """

        WriteTextCursor.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WordCursorPartial.__init__(self, component, None)  # type: ignore
        StylePartial.__init__(self, component=component)
