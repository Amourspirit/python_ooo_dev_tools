from __future__ import annotations
from typing import Any, List, Tuple, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertySet
from ooo.dyn.style.numbering_type import NumberingTypeEnum
from ooo.dyn.text.page_number_type import PageNumberType

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextRange
    from com.sun.star.text import XText
    from ooodev.proto.style_obj import StyleT

from ooodev.office import write as mWrite
from ooodev.utils import lo as mLo
from ooodev.adapter.text.text_document_comp import TextDocumentComp
from ooodev.adapter.container.name_access_comp import NameAccessComp
from ooodev.adapter.text.textfield.page_count_comp import PageCountComp
from ooodev.adapter.text.textfield.page_number_comp import PageNumberComp
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.event_singleton import _Events
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.type_var import PathOrStr
from . import write_text_cursor as mWriteTextCursor
from . import write_character_style as mWriteCharacterStyle
from . import write_paragraph_style as mWriteParagraphStyle
from . import write_text_content as mWriteTextContent
from . import write_text_view_cursor as mWriteTextViewCursor


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
    def get_cursor(self) -> mWriteTextCursor.WriteTextCursor:
        """
        Gets text cursor from the current document.

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    @overload
    def get_cursor(self, *, cursor_obj: Any) -> mWriteTextCursor.WriteTextCursor:
        """
        Gets text cursor

        Args:
            cursor_obj (Any): Text Document or Text Cursor

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    @overload
    def get_cursor(self, *, rng: XTextRange, txt: XText) -> mWriteTextCursor.WriteTextCursor:
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
    def get_cursor(self, *, rng: XTextRange) -> mWriteTextCursor.WriteTextCursor:
        """
        Gets text cursor

        Args:
            rng (XTextRange): Text Range instance

        Returns:
            WriteTextCursor: Cursor
        """
        ...

    def get_cursor(self, **kwargs) -> mWriteTextCursor.WriteTextCursor:
        """Returns the cursor of the document."""
        if not kwargs:
            return mWriteTextCursor.WriteTextCursor(self, self.component.getText().createTextCursor())
        if "text_doc" in kwargs:
            kwargs["text_doc"] = self.component
        return mWriteTextCursor.WriteTextCursor(self, mWrite.Write.get_cursor(**kwargs))

    # endregion get_cursor()

    def close_doc(self) -> bool:
        """
        Closes text document

        Returns:
            bool: False if DOC_CLOSING event is canceled, Other

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc``.

        Attention:
            :py:meth:`Lo.close <.utils.lo.Lo.close>` method is called along with any of its events.
        """

        cargs = CancelEventArgs(self.close_doc.__qualname__)
        cargs.event_data = {"text_doc": self.component}
        _Events().trigger(WriteNamedEvent.DOC_CLOSING, cargs)
        if cargs.cancel:
            return False
        result = mLo.Lo.close(self.component)  # type: ignore
        _Events().trigger(WriteNamedEvent.DOC_CLOSED, EventArgs.from_args(cargs))
        return result

    def create_style_char(
        self, style_name: str, styles: Sequence[StyleT] | None = None
    ) -> mWriteCharacterStyle.WriteCharacterStyle:
        """
        Creates a character style and adds it to document character styles.

        Args:
            text_doc (XTextDocument): Text Document
            style_name (str): The name of the character style.
            styles (Sequence[StyleT], optional): One or more styles to apply.

        Returns:
            WriteCharacterStyle: Newly created style
        """
        if styles is None:
            styles = ()
        result = mWrite.Write.create_style_char(self.component, style_name, styles)
        return mWriteCharacterStyle.WriteCharacterStyle(self, result)

    def create_style_para(
        self, style_name: str, styles: Sequence[StyleT] | None = None
    ) -> mWriteParagraphStyle.WriteParagraphStyle:
        """
        Creates a paragraph style and adds it to document paragraph styles.

        Args:
            style_name (str): The name of the paragraph style.
            styles (Sequence[StyleT], optional): One or more styles to apply.

        Returns:
            WriteParagraphStyle: Newly created style
        """
        if styles is None:
            styles = ()
        result = mWrite.Write.create_style_para(self.component, style_name, styles)
        return mWriteParagraphStyle.WriteParagraphStyle(self, result)

    def find_bookmark(self, bm_name: str) -> mWriteTextContent.WriteTextContent | None:
        """
        Finds a bookmark

        Args:
            bm_name (str): Bookmark name

        Returns:
            WriteTextContent | None: Bookmark if found; Otherwise, None
        """
        result = mWrite.Write.find_bookmark(self.component, bm_name)
        if result is None:
            return None
        return mWriteTextContent.WriteTextContent(self, result)

    def get_view_cursor(self) -> mWriteTextViewCursor.WriteTextViewCursor:
        """
        Gets document view cursor.

        Describes a cursor in a text document's view.

        Raises:
            ViewCursorError: If Unable to get cursor

        Returns:
            WriteTextViewCursor: Text View Cursor

        See Also:
            `LibreOffice API XTextViewCursor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextViewCursor.html>`_
        """
        cursor = mWrite.Write.get_view_cursor(self.component)
        return mWriteTextViewCursor.WriteTextViewCursor(self, cursor)

    def get_doc_settings(self) -> XPropertySet:
        """
        Gets Text Document Settings

        Returns:
            XPropertySet: Settings

        See Also:
            `API DocumentSettings Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1DocumentSettings.html>`__

        .. versionadded:: 0.9.7
        """
        return mLo.Lo.create_instance_msf(XPropertySet, "com.sun.star.text.DocumentSettings", raise_err=True)

    def get_graphic_links(self) -> NameAccessComp | None:
        """
        Gets graphic links

        Args:
            doc (XComponent): Document

        Raises:
            MissingInterfaceError: if doc does not implement ``XTextGraphicObjectsSupplier`` interface

        Returns:
            NameAccessComp | None: Graphic Links on success, Otherwise, None
        """
        result = mWrite.Write.get_graphic_links(self.component)
        if result is None:
            return None
        return NameAccessComp(result)

    def get_num_of_pages(self) -> int:
        """
        Gets document page count

        Returns:
            int: page count
        """
        return mWrite.Write.get_num_of_pages(self.component)

    def get_page_count_field(self, numbering_type: NumberingTypeEnum = NumberingTypeEnum.ARABIC) -> PageCountComp:
        """
        Gets page count field

        Returns:
            PageCountComp: Page Count Field
        """
        return PageCountComp(mWrite.Write.get_page_count(numbering_type=numbering_type))

    def get_page_number_field(
        self,
        numbering_type: NumberingTypeEnum = NumberingTypeEnum.ARABIC,
        sub_type: PageNumberType = PageNumberType.CURRENT,
    ) -> PageNumberComp:
        """
        Gets Arabic style number showing current page value

        Args:
            numbering_type (NumberingTypeEnum, optional): Numbering Type. Defaults to ``NumberingTypeEnum.ARABIC``.
            sub_type (PageNumberType, optional): Page Number Type. Defaults to ``PageNumberType.CURRENT``.

        Returns:
            PageNumberComp: Page Number Field
        """
        return PageNumberComp(mWrite.Write.get_page_number(numbering_type=numbering_type, sub_type=sub_type))
