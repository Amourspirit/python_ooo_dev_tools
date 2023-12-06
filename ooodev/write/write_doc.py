from __future__ import annotations
from typing import Any, List, Tuple, overload, Sequence, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextRange
    from com.sun.star.text import XText

from ooodev.office import write as mWrite
from ooodev.adapter.text.text_document_comp import TextDocumentComp
from ooodev.events.event_singleton import _Events
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.type_var import PathOrStr
from .write_text_cursor import WriteTextCursor


class WriteDoc(TextDocumentComp, QiPartial, PropPartial):
    """A class to represent a Write document."""

    def __init__(self, doc: XTextDocument) -> None:
        """
        Constructor

        Args:
            doc (XTextDocument): A UNO object that supports ``com.sun.star.sheet.TextDocument`` service.
        """
        TextDocumentComp.__init__(self, doc)  # type: ignore
        QiPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)

    # region get_cursor()
    @overload
    def get_cursor(self) -> WriteTextCursor:
        """
        Gets text cursor from the current document.

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    @overload
    def get_cursor(self, *, cursor_obj: Any) -> WriteTextCursor:
        """
        Gets text cursor

        Args:
            cursor_obj (Any): Text Document or Text Cursor

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    @overload
    def get_cursor(self, *, rng: XTextRange, txt: XText) -> WriteTextCursor:
        """
        Gets text cursor

        Args:
            rng (XTextRange): Text Range Instance
            txt (XText): Text Instance

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    @overload
    def get_cursor(self, *, rng: XTextRange) -> WriteTextCursor:
        """
        Gets text cursor

        Args:
            rng (XTextRange): Text Range instance

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    def get_cursor(self, **kwargs) -> WriteTextCursor:
        """Returns the cursor of the document."""
        if not kwargs:
            return WriteTextCursor(self, self.component.getText().createTextCursor())
        if "text_doc" in kwargs:
            kwargs["text_doc"] = self.component
        return WriteTextCursor(self, mWrite.Write.get_cursor(**kwargs))

    # endregion get_cursor()
