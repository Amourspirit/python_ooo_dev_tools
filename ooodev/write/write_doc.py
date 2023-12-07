from __future__ import annotations
from typing import Any, List, Tuple, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertySet
from ooo.dyn.style.numbering_type import NumberingTypeEnum
from ooo.dyn.text.page_number_type import PageNumberType

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextRange
    from com.sun.star.text import XText
    from ooo.dyn.view.paper_format import PaperFormat
    from ooodev.proto.style_obj import StyleT

from ooodev.office import write as mWrite
from ooodev.utils import selection as mSelection
from ooodev.utils import lo as mLo
from ooodev.utils.data_type.size import Size
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
from . import write_draw_page as mWriteDrawPage
from . import write_text_range as mWriteTextRange


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

    def compare_cursor_ends(self, c1: XTextRange, c2: XTextRange) -> mSelection.Selection.CompareEnum:
        """
        Compares two cursors ranges end positions

        Args:
            c1 (XTextRange): first cursor range
            c2 (XTextRange): second cursor range

        Raises:
            Exception: if comparison fails

        Returns:
            CompareEnum: Compare result.
            :py:attr:`.CompareEnum.BEFORE` if ``c1`` end position is before ``c2`` end position.
            :py:attr:`.CompareEnum.EQUAL` if ``c1`` end position is equal to ``c2`` end position.
            :py:attr:`.CompareEnum.AFTER` if ``c1`` end position is after ``c2`` end position.
        """
        return mSelection.Selection.compare_cursor_ends(c1, c2)

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
        cursor = mSelection.Selection.get_view_cursor(self.component)
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

    def get_left_cursor(self, rng: XTextRange) -> mWriteTextCursor.WriteTextCursor:
        """
        Creates a new TextCursor with position left that can travel right.

        Args:
            rng (XTextRange): Text Range.

        Returns:
            WriteTextCursor: a new instance of a TextCursor which is located at the specified
            TextRange to travel in the given text context.
        """
        result = mSelection.Selection.get_left_cursor(rng, self.component)
        return mWriteTextCursor.WriteTextCursor(self, result)

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

    def get_page_size(self) -> Size:
        """
        Get page size in ``1/100 mm`` units.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            ~ooodev.utils.data_type.size.Size: Page Size in ``1/100 mm`` units.
        """
        return mWrite.Write.get_page_size(self.component)

    def get_page_text_size(self) -> Size:
        """
        Get page text size in ``1/100 mm`` units.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            ~ooodev.utils.data_type.size.Size: Page text Size in ``1/100 mm`` units.
        """
        return mWrite.Write.get_page_text_size(self.component)

    def get_page_text_width(self) -> int:
        """
        Get the width of the page's text area in ``1/100 mm`` units.

        Returns:
            int: Page Width in ``1/100 mm`` units on success; Otherwise 0
        """
        return mWrite.Write.get_page_text_width(self.component)

    def get_selected_text_range(self) -> mWriteTextRange.WriteTextRange | None:
        """
        Gets the text range for current selection

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If unable to obtain required interface.

        Returns:
            WriteTextRange | None: If no selection is made then None is returned; Otherwise, Text Range.

        Note:
            Writer must be visible for this method or ``None`` is returned.
        """
        result = mSelection.Selection.get_selected_text_range(self.component)
        if result is None:
            return None
        return mWriteTextRange.WriteTextRange(self, result)

    def get_selected_text_str(self) -> str:
        """
        Gets the first selection text for Document

        Returns:
            str: Selected text or empty string.

        Note:
            Writer must be visible for this method or empty string is returned.
        """
        return mSelection.Selection.get_selected_text_str(self.component)

    def get_draw_page(self) -> mWriteDrawPage.WriteDrawPage:
        """
        Gets draw page.

        Returns:
            WriteDrawPage: Draw Page
        """
        draw_page = mWrite.Write.get_draw_page(self.component)
        return mWriteDrawPage.WriteDrawPage(self, draw_page)

    def get_right_cursor(self, rng: XTextRange) -> mWriteTextCursor.WriteTextCursor:
        """
        Creates a new TextCursor with position right that can travel left.

        Args:
            rng (XTextRange): Text Range.

        Returns:
            WriteTextCursor: a new instance of a TextCursor which is located at the specified
            TextRange to travel in the given text context.
        """
        result = mSelection.Selection.get_right_cursor(rng, self.component)
        return mWriteTextCursor.WriteTextCursor(self, result)

    def get_text_frames(self) -> NameAccessComp | None:
        """
        Gets document Text Frames.

        Args:
            doc (XComponent): Document

        Raises:
            MissingInterfaceError: if doc does not implement ``XTextFramesSupplier`` interface

        Returns:
            NameAccessComp | None: Text Frames on success, Otherwise, None
        """
        result = mWrite.Write.get_text_frames(self.component)
        if result is None:
            return None
        return NameAccessComp(result)

    def get_text_graphics(self) -> List[XGraphic]:
        """
        Gets text graphics.

        Raises:
            Exception: If unable to get text graphics

        Returns:
            List[XGraphic]: Text Graphics

        Note:
            If there is error getting a graphic link then it is ignored
            and not added to the return value.
        """
        return mWrite.Write.get_text_graphics(self.component)

    def is_anything_selected(self) -> bool:
        """
        Determine if anything is selected.

        If Write document is not visible this method returns false.

        Returns:
            bool: True if anything in the document is selected: Otherwise, False

        Note:
            Writer must be visible for this method or ``False`` is always returned.
        """
        return mSelection.Selection.is_anything_selected(self.component)

    def select_next_word(self) -> bool:
        """
        Select the word right from the current cursor position.

        Returns:
            bool: True if go to next word succeeds; Otherwise, False.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.WORD_SELECTING` :eventref:`write_word_selecting`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.WORD_SELECTED` :eventref:`write_word_selected`

        Note:
            The method returning ``True`` does not necessarily mean that the cursor is located at
            the next word, or any word at all! This may happen for example if it travels over empty paragraphs.

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc``.
        """
        return mSelection.Selection.select_next_word(self.component)

    def save_doc(self, fnm: PathOrStr) -> bool:
        """
        Saves text document

        Args:
            fnm (PathOrStr): Path to save as

        Returns:
            bool: True if doc is saved; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc`` and ``fnm``.

        Attention:
            :py:meth:`Lo.save_doc <.utils.lo.Lo.save_doc>` method is called along with any of its events.
        """
        return mWrite.Write.save_doc(self.component, fnm)

    def set_a4_page_format(
        self,
    ) -> bool:
        """
        Set Page Format to A4

        Returns:
            bool: ``True`` if page format is set; Otherwise, ``False``

        See Also:
            :py:meth:`~.write_doc.WriteDoc.set_page_format`

        Attention:
            :py:meth:`~.write.Write.set_page_format` method is called along with any of its events.
        """
        return mWrite.Write.set_a4_page_format(self.component)

    def set_footer(self, text: str, styles: Sequence[StyleT] | None = None) -> None:
        """
        Modify the footer via the page style for the document.
        Put the text on the right hand side in the header in
        a general font of 10pt.

        Args:
            text (str): Header Text
            styles (Sequence[StyleT]): Styles to apply to the text.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to set header text

        See Also:
            :py:meth:`~.write.Write.set_header`

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_general_name`
        """
        mWrite.Write.set_footer(self.component, text, styles)

    def set_header(self, text: str, styles: Sequence[StyleT] | None = None) -> None:
        """
        Modify the header via the page style for the document.
        Put the text on the right hand side in the header in
        a general font of 10pt.

        Args:
            text (str): Header Text
            styles (Sequence[StyleT]): Styles to apply to the text.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to set header text

        See Also:
            :py:meth:`~.write.Write.set_footer`

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_general_name`
        """
        mWrite.Write.set_header(self.component, text, styles)

    def set_page_format(self, paper_format: PaperFormat) -> bool:
        """
        Set Page Format

        Args:
            paper_format (~com.sun.star.view.PaperFormat): Paper Format.

        Raises:
            MissingInterfaceError: If ``text_doc`` does not implement ``XPrintable`` interface

        Returns:
            bool: ``True`` if page format is set; Otherwise, ``False``

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.PAGE_FORMAT_SETTING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.PAGE_FORMAT_SET` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``fnm``.

        See Also:
            - :py:meth:`~.write_doc.WriteDoc.set_a4_page_format`
        """
        return mWrite.Write.set_page_format(self.component, paper_format)

    def set_page_numbers(self) -> PageNumberComp:
        """
        Modify the footer via the page style for the document.
        Put page number & count in the center of the footer in Times New Roman, 12pt

        Raises:
            PropertiesError: If unable to get properties
            Exception: If Unable to set page numbers

        Returns:
            PageNumberComp: Page Number Field

        .. versionchanged:: 0.16.0
            Returns ``XTextField``.
        """
        result = mWrite.Write.set_page_numbers(self.component)
        return PageNumberComp(result)
