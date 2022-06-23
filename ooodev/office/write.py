# coding: utf-8
# Python conversion of Write.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Iterable, List, overload
import re
import os
import uno

from ..exceptions import ex as mEx
from ..utils import lo as mLo
from ..utils import info as mInfo
from ..utils import file_io as mFileIO
from ..utils import props as mProps
from ..utils.gen_util import TableHelper
from ..utils.color import CommonColor, Color
from ..utils.type_var import PathOrStr, Table, DocOrCursor
from ..utils import images_lo as mImgLo
from ..utils import selection as mSel
from ..mock import mock_g


if not mock_g.DOCS_BUILDING:
    # not importing for doc building just result in short import name for
    # args that use these.
    # this is also true becuase docs/conf.py ignores com import for autodoc
    from com.sun.star.awt import FontWeight
    from com.sun.star.beans import XPropertySet
    from com.sun.star.container import XEnumerationAccess
    from com.sun.star.container import XNamed
    from com.sun.star.document import XDocumentInsertable
    from com.sun.star.document import XEmbeddedObjectSupplier2
    from com.sun.star.drawing import XDrawPageSupplier
    from com.sun.star.drawing import XShape
    from com.sun.star.lang import Locale  # struct class
    from com.sun.star.lang import XComponent
    from com.sun.star.lang import XServiceInfo
    from com.sun.star.frame import XModel
    from com.sun.star.linguistic2 import XConversionDictionaryList
    from com.sun.star.linguistic2 import XLanguageGuessing
    from com.sun.star.linguistic2 import XLinguProperties
    from com.sun.star.linguistic2 import XLinguServiceManager
    from com.sun.star.linguistic2 import XProofreader
    from com.sun.star.linguistic2 import XSearchableDictionaryList
    from com.sun.star.style import NumberingType  # const
    from com.sun.star.table import BorderLine  # struct
    from com.sun.star.text import HoriOrientation
    from com.sun.star.text import VertOrientation
    from com.sun.star.text import XBookmarksSupplier
    from com.sun.star.text import XPageCursor
    from com.sun.star.text import XParagraphCursor
    from com.sun.star.text import XText
    from com.sun.star.text import XTextContent
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextGraphicObjectsSupplier
    from com.sun.star.text import XTextField
    from com.sun.star.text import XTextFrame
    from com.sun.star.text import XTextRange
    from com.sun.star.text import XTextTable
    from com.sun.star.text import XTextViewCursor
    from com.sun.star.text import XTextViewCursorSupplier
    from com.sun.star.uno import Exception as UnoException
    from com.sun.star.util import XCloseable
    from com.sun.star.view import XPrintable

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XEnumeration
    from com.sun.star.container import XNameAccess
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XFrame
    from com.sun.star.graphic import XGraphic
    from com.sun.star.linguistic2 import SingleProofreadingError
    from com.sun.star.linguistic2 import XLinguServiceManager2
    from com.sun.star.linguistic2 import XSpellChecker
    from com.sun.star.linguistic2 import XThesaurus
    from com.sun.star.text import XTextCursor


from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.size import Size  # struct
from ooo.dyn.linguistic2.dictionary_type import DictionaryType as OooDictionaryType
from ooo.dyn.style.break_type import BreakType
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust
from ooo.dyn.text.control_character import ControlCharacterEnum
from ooo.dyn.text.page_number_type import PageNumberType
from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType
from ooo.dyn.view.paper_format import PaperFormat as OooPaperFormat
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust

# endregion Imports


class Write(mSel.Selection):
    ControlCharacter = ControlCharacterEnum  # UnoControlCharacter
    # region -------------- Enums --------------------------------------

    DictionaryType = OooDictionaryType

    PaperFormat = OooPaperFormat

    ParagraphAdjust = ParagraphAdjust

    # endregion ----------- Enums --------------------------------------

    # region ------------- doc / open / close /create/ etc -------------
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XTextDocument:
        """
        Opens a Text (Writer) document.

        Args:
            fnm (PathOrStr): Spreadsheet file to open
            loader (XComponentLoader): Component loader

        Raises:
            Exception: If Document is Null
            Exception: If Not a Text Doucment
            MissingInterfaceError: If unable to obtain XTextDocument interface

        Returns:
            XTextDocument: Text Document
        """
        doc = mLo.Lo.open_doc(fnm=fnm, loader=loader)
        if doc is None:
            raise Exception("Document is null")
        if not cls.is_text(doc):
            print(f"Not a text document; closing '{fnm}'")
            mLo.Lo.close_doc(doc)
            raise Exception("Not a text document")
        text_doc = mLo.Lo.qi(XTextDocument, doc)
        if text_doc is None:
            print(f"Not a text document; closing '{fnm}'")
            mLo.Lo.close_doc(doc)
            raise mEx.MissingInterfaceError(XTextDocument)
        return text_doc

    @staticmethod
    def is_text(doc: XComponent) -> bool:
        """
        Gets if doc is an actual Writer document

        Args:
            doc (XComponent): Document Component

        Returns:
            bool: True if doc is Writer Document; Otherwise, False
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.WRITER)

    @classmethod
    def get_text_doc(cls, doc: XComponent) -> XTextDocument:
        """
        Gets a writer document

        When using this method in a macro the ``Lo.get_document()`` value should be passed as ``doc`` arg.

        Args:
            doc (XComponent): Component to get writer document from

        Raises:
            TypeError: doc is None
            MissingInterfaceError: If doc does not implement XTextDocument interface

        Returns:
            XTextDocument: Writer document
        """
        if doc is None:
            raise TypeError("Document is null")

        text_doc = mLo.Lo.qi(XTextDocument, doc, True)
        return text_doc

    @staticmethod
    def create_doc(loader: XComponentLoader) -> XTextDocument:
        """
        Creates a new Writer Text Document

        Args:
            loader (XComponentLoader): Component Loader

        Returns:
            XTextDocument: Text Document
        """
        return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.WRITER, loader=loader)

    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader) -> XTextDocument:
        """
        Create a new Writer Text Document from a template

        Args:
            template_path (PathOrStr): Path to Template
            loader (XComponentLoader): Component Loader

        Raises:
            MissingInterfaceError: If Unable to obtain XTextDocument interface

        Returns:
            XTextDocument: Text Document
        """
        doc = mLo.Lo.create_doc_from_template(template_path=template_path, loader=loader)
        xdoc = mLo.Lo.qi(XTextDocument, doc, True)
        return xdoc

    @staticmethod
    def close_doc(text_doc: XTextDocument) -> None:
        """
        Closes text document

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If unable to obtain XCloseable from text_doc
        """
        closable = mLo.Lo.qi(XCloseable, text_doc, True)
        mLo.Lo.close(closable)

    @staticmethod
    def save_doc(text_doc: XTextDocument, fnm: PathOrStr) -> None:
        """
        Saves text document

        Args:
            text_doc (XTextDocument): Text Document
            fnm (PathOrStr): Path to save as

        Raises:
            MissingInterfaceError: If text_doc does not implement XComponent interface
        """
        doc = mLo.Lo.qi(XComponent, text_doc, True)
        mLo.Lo.save_doc(doc=doc, fnm=fnm)

    @classmethod
    def open_flat_doc_using_text_template(
        cls, fnm: PathOrStr, template_path: PathOrStr, loader: XComponentLoader
    ) -> XTextDocument:
        """
        Open a new text document applying the template as formatting to the flat XML file

        Args:
            fnm (PathOrStr): path to file
            template_path (PathOrStr): Path to template file (ott)
            loader (XComponentLoader): Component Loader

        Raises:
            UnOpenableError: If fnm is not openable
            ValueError: If template_path is not ott file
            MissingInterfaceError: If template_path document does not implement XTextDocument interface
            ValueError: If unable to obtain cursor object
            Exception: Any other errors

        Returns:
            XTextDocument | None: Text Document
        """
        if fnm is None:
            print("Filename is null")
            return None

        open_file_url = None
        if not mFileIO.FileIO.is_openable(fnm):
            if mLo.Lo.is_url(fnm):
                print(f"Treating filename as a URL: '{fnm}'")
                open_file_url = str(fnm)
            else:
                raise mEx.UnOpenableError(fnm)
        else:
            open_file_url = mFileIO.FileIO.fnm_to_url(fnm)

        template_ext = mInfo.Info.get_ext(template_path)
        if template_ext != "ott":
            raise ValueError(f"Can only apply a text template as formatting. Not an ott file: {template_path}")

        doc = mLo.Lo.create_doc_from_template(template_path=template_path, loader=loader)
        text_doc = mLo.Lo.qi(XTextDocument, doc)
        if text_doc is None:
            raise mEx.MissingInterfaceError(
                XTextDocument, f"Template is not a text document. Missing: {XTextDocument.__pyunointerface__}"
            )

        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            raise ValueError(f"Unable to get cursor: '{fnm}'")

        try:
            cursor.gotoEnd(True)
            di = mLo.Lo.qi(XDocumentInsertable, cursor)
            # XDocumentInsertable only works with text files
            if di is None:
                print("Document inserter could not be created")
            else:
                di.insertDocumentFromURL(open_file_url, tuple())
                # Props.makeProps("FilterName", "OpenDocument Text Flat XML"))
                # these props do not work
        except Exception as e:
            raise Exception("Could not insert document") from e
        return text_doc

    # endregion ---------- doc / open / close /create/ etc -------------

    # region ------------- page methods --------------------------------
    @classmethod
    def get_page_cursor(cls, text_doc: XTextDocument) -> XPageCursor:
        """
        Get Page curosor

        Makes it possible to perform cursor movements between pages.

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PageCursorError: If Unable to get cursor

        Returns:
            XPageCursor: Page Cursor

        See Also:
            `LibreOffice API XPageCursor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XPageCursor.html>`_
        """
        try:
            view_cursor = cls.get_view_cursor(text_doc)
            page_cursor = mLo.Lo.qi(XPageCursor, view_cursor, True)
            return page_cursor
        except Exception as e:
            raise mEx.PageCursorError(str(e)) from e

    @staticmethod
    def get_current_page(tv_cursor: XTextViewCursor) -> int:
        """
        Gets the current page

        Args:
            tv_cursor (XTextViewCursor): Text view Cursor

        Returns:
            int: Page number if present; Otherwise, -1
        """
        page_cursor = mLo.Lo.qi(XPageCursor, tv_cursor)
        if page_cursor is None:
            print("Could not create a page cursor")
            return -1
        return page_cursor.getPage()

    # @classmethod
    # def get_coord(cls, text_doc: XTextDocument) -> Point:
    #     # see section 7.17 Useful Macro Information For OpenOffice By Andrew Pitonyak.pdf

    #     tvc = cls.get_view_cursor(text_doc)
    #     if not mInfo.Info.support_service(tvc, "com.sun.star.style.ParagraphStyle"):
    #         mEx.NotSupportedServiceError("com.sun.star.style.ParagraphStyle")
    #     ps = cast("ParagraphStyle", tvc)
    #     props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
    #     if props is None:
    #         raise mEx.PropertiesError("Could not access the standard page style")
    #     lHeight = int(props.getPropertyValue("Width"))
    #     lWidth = int(props.getPropertyValue("Height"))
    #     top_margin = int(props.getPropertyValue("TopMargin"))
    #     left_margin = int(props.getPropertyValue("LeftMargin"))
    #     bottom_margin = int(props.getPropertyValue("BottomMargin"))

    #     char_height = int(props.getPropertyValue("CharHeight"))

    #     dCharHeight = char_height / 72.0
    #     page_cursor = mLo.Lo.qi(XPageCursor, tvc)
    #     if page_cursor is None:
    #         raise mEx.MissingInterfaceError(XPageCursor)
    #     iCurPage = page_cursor.getPage()
    #     v = tvc.getPosition()
    #     dYCursor = (v.Y + top_margin)/2540.0 + dCharHeight / 2
    #     dXCursor = (v.X + left_margin)/2540.0
    #     dXRight = (lWidth - v.X - left_margin)/2540.0
    #     dYBottom = (lHeight - v.Y - top_margin)/2540.0 - dCharHeight / 2

    #     dBottomMargin = bottom_margin / 2540.0
    #     dLeftMargin = left_margin / 2540.0

    @staticmethod
    def get_coord_str(tv_cursor: XTextViewCursor) -> str:
        """
        Gets coordinates for cursor in format such as ``"10, 10"``

        Args:
            tv_cursor (XTextViewCursor): Text View Cursor

        Returns:
            str: coordinates as string
        """
        pos = tv_cursor.getPosition()
        return f"({pos.X}, {pos.Y})"

    @staticmethod
    def get_page_count(text_doc: XTextDocument) -> int:
        """
        Gets document page count

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If text_doc does not implement XModel interface

        Returns:
            int: page count
        """
        model = mLo.Lo.qi(XModel, text_doc, True)
        xcontroller = model.getCurrentController()
        return int(mProps.Props.get_property(xcontroller, "PageCount"))

    # endregion ---------- page methods --------------------------------

    # region ------------- text writing methods ------------------------

    # region    append()
    @classmethod
    def _append_text(cls, cursor: XTextCursor, text: str) -> None:
        cursor.setString(text)
        cursor.gotoEnd(False)

    @classmethod
    def _append_ctl_char(cls, cursor: XTextCursor, ctl_char: int) -> None:
        xtext = cursor.getText()
        xtext.insertControlCharacter(cursor, ctl_char, False)
        cursor.gotoEnd(False)

    @classmethod
    def _append_text_content(cls, cursor: XTextCursor, text_content: XTextContent) -> None:
        xtext = cursor.getText()
        xtext.insertTextContent(cursor, text_content, False)
        cursor.gotoEnd(False)

    @overload
    @staticmethod
    def append(cursor: XTextCursor, text: str) -> None:
        """
        Appends text to text cursor

        Args:
            cursor (XTextCursor): Text Cursor
            text (str): Text to append
        """
        ...

    @overload
    @staticmethod
    def append(cursor: XTextCursor, ctl_char: ControlCharacter) -> None:
        """
        Appents a control character (like a paragraph break or a hard space) into the text.

        Args:
            cursor (XTextCursor): Text Cursor
            ctl_char (Write.ControlCharacter): Control Char
        """
        ...

    @overload
    @staticmethod
    def append(cursor: XTextCursor, text_content: XTextContent) -> None:
        """
        Appends a content, such as a text table, text frame or text field.

        Args:
            cursor (XTextCursor): Text Cursor
            text_content (XTextContent): Text Content
        """
        ...

    @classmethod
    def append(cls, *args, **kwargs) -> None:
        """
        Append content to cursor

        Args:
            cursor (XTextCursor): Text Cursor
            text (str): Text to append
            ctl_char (int): Control Char (like a paragraph break or a hard space)
            text_content (XTextContent): Text content, such as a text table, text frame or text field.

        See Also:
            `API ControlCharacter <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1ControlCharacter.html>`_
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cursor", "text", "ctl_char", "text_content")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("append() got an unexpected keyword argument")
            ka[1] = kwargs.get("cursor", None)
            keys = ("text", "ctl_char", "text_content")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("append() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if isinstance(kargs[2], str):
            cls._append_text(cursor=kargs[1], text=kargs[2])
        elif isinstance(kargs[2], int):
            cls._append_ctl_char(cursor=kargs[1], ctl_char=kargs[2])
        else:
            cls._append_text_content(cursor=kargs[1], text_content=kargs[2])

    # endregion append()

    @classmethod
    def append_line(cls, cursor: XTextCursor, text: str | None = None) -> None:
        """
        Appends a new Line

        Args:
            cursor (XTextCursor): Text Cursor
            text (str, optional): text to append before new line is inserted.
        """
        if text is not None:
            cls._append_text(cursor=cursor, text=text)
        cls._append_ctl_char(cursor=cursor, ctl_char=Write.ControlCharacter.LINE_BREAK)

    @classmethod
    def append_date_time(cls, cursor: XTextCursor) -> None:
        """
        Append two DateTime fields, one for the date, one for the time

        Args:
            cursor (XTextCursor): Text Cursor

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
        """
        dt_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
        mProps.Props.set_property(dt_field, "IsDate", True)  # so date is reported
        xtext_content = mLo.Lo.qi(XTextContent, dt_field, True)
        cls._append_text_content(cursor, xtext_content)
        cls.append(cursor, "; ")

        dt_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
        mProps.Props.set_property(dt_field, "IsDate", False)  # so time is reported
        xtext_content = mLo.Lo.qi(XTextContent, dt_field, True)
        cls._append_text_content(cursor, xtext_content)

    @classmethod
    def append_para(cls, cursor: XTextCursor, text: str) -> None:
        """
        Appends text and then a paragraph break.

        Args:
            cursor (XTextCursor): Text Cursor
            text (str): Text to append
        """
        cls._append_text(cursor=cursor, text=text)
        cls._append_ctl_char(cursor=cursor, ctl_char=Write.ControlCharacter.PARAGRAPH_BREAK)

    @classmethod
    def end_line(cls, cursor: XTextCursor) -> None:
        """
        Inserts a line break

        Args:
            cursor (XTextCursor): Text Cursor
        """
        cls._append_ctl_char(cursor=cursor, ctl_char=Write.ControlCharacter.LINE_BREAK)

    @classmethod
    def end_paragraph(cls, cursor: XTextCursor) -> None:
        """
        Inserts a paragraph break

        Args:
            cursor (XTextCursor): Text Cursor
        """
        cls._append_ctl_char(cursor=cursor, ctl_char=Write.ControlCharacter.PARAGRAPH_BREAK)

    @classmethod
    def page_break(cls, cursor: XTextCursor) -> None:
        """
        Inserts a page break

        Args:
            cursor (XTextCursor): Text Cursor
        """
        mProps.Props.set_property(cursor, "BreakType", BreakType.PAGE_AFTER)
        cls.end_paragraph(cursor)

    @classmethod
    def column_break(cls, cursor: XTextCursor) -> None:
        """
        Inserts a column break

        Args:
            cursor (XTextCursor): Text Cursor
        """
        mProps.Props.set_property(cursor, "BreakType", BreakType.COLUMN_AFTER)
        cls.end_paragraph(cursor)

    @classmethod
    def insert_para(cls, cursor: XTextCursor, para: str, para_style: str) -> None:
        """
        Inserts a paragraph with a style applied

        Args:
            cursor (XTextCursor): Text Cursor
            para (str): Paragraph text
            para_style (str): Style such as 'Heading 1'
        """
        xtext = cursor.getText()
        xtext.insertString(cursor, para, False)
        xtext.insertControlCharacter(cursor, Write.ControlCharacter.PARAGRAPH_BREAK, False)
        cls.style_prev_paragraph(cursor, para_style)

    # endregion ---------- text writing methods ------------------------

    # region ------------- extract text from document ------------------

    @staticmethod
    def get_all_text(cursor: XTextCursor) -> str:
        """
        Gets the text part of the document

        Args:
            cursor (XTextCursor): Text Cursor

        Returns:
            str: text
        """
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        text = cursor.getString()
        cursor.gotoEnd(False)  # to deselect everything
        return text

    @staticmethod
    def get_enumeration(obj: object) -> XEnumeration:
        """
        Gets Enumeration access from obj

        Used to enumerate objects in a container which contains objects.

        Args:
            obj (object): object that implements XEnumerationAccess or XTextDocument.

        Raises:
            MissingInterfaceError: if obj does not implement XEnumerationAccess interface

        Returns:
            XEnumeration: Enumerator
        """
        enum_access = mLo.Lo.qi(XEnumerationAccess, obj)
        if enum_access is None:
            # try for XTextDocument
            try:
                xtext = obj.getText()
                if xtext is not None:
                    enum_access = mLo.Lo.qi(XEnumerationAccess, xtext)
            except AttributeError:
                pass
        if enum_access is None:
            raise mEx.MissingInterfaceError(XEnumerationAccess)
        return enum_access.createEnumeration()

    # endregion ---------- extract text from document ------------------

    # region ------------- text cursor property methods ----------------

    @classmethod
    def style_left_bold(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles bold from current cursor postiion left by pos amount.

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Number of postiions to go left
        """
        cls.style_left(cursor, pos, "CharWeight", FontWeight.BOLD)

    @classmethod
    def style_left_italic(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles italic from current cursor postiion left by pos amount.

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Number of postiions to go left
        """
        cls.style_left(cursor, pos, "CharPosture", FontSlant.ITALIC)

    @classmethod
    def style_left_color(cls, cursor: XTextCursor, pos: int, color: Color) -> None:
        """
        Styles color from current cursor postiion left by pos amount.

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Number of postiions to go left
            color (Color): RGB color as int to apply

        See Also:
            :py:class:`~.utils.color.CommonColor`
        """
        cls.style_left(cursor, pos, "CharColor", color)

    @classmethod
    def style_left_code(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles using a Mono font from current cursor postiion left by pos amount.
        Font Char Height is set to ``10``

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Number of postiions to go left

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_mono_name`
        """
        cls.style_left(cursor, pos, "CharFontName", mInfo.Info.get_font_mono_name())
        cls.style_left(cursor, pos, "CharHeight", 10)

    @classmethod
    def style_left(cls, cursor: XTextCursor, pos: int, prop_name: str, prop_val: object) -> None:
        """
        Styles left. From current cursor postiion to left by pos amount.

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Postiions to style left
            prop_name (str): Property Name such as 'CharHeight
            prop_val (object): Property Value such as 10
        """
        old_val = mProps.Props.get_property(cursor, prop_name)

        curr_pos = mSel.Selection.get_position(cursor)
        cursor.goLeft(curr_pos - pos, True)
        mProps.Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)

        cursor.goRight(curr_pos - pos, False)
        mProps.Props.set_property(prop_set=cursor, name=prop_name, value=old_val)

    @classmethod
    def dispatch_cmd_left(
        cls,
        vcursor: XTextViewCursor,
        pos: int,
        cmd: str,
        props: Iterable[PropertyValue] = None,
        frame: XFrame = None,
        toggle: bool = False,
    ) -> None:
        """
        Dispatches a command and applies it to selection based upon position

        Args:
            vcursor (XTextViewCursor): Text View Cursor
            pos (int): Positions left to apply dispacch command
            cmd (str): Dispatach command such as 'DefaultNumbering'
            props (Iterable[PropertyValue], optional): properties for dispatch
            frame (XFrame, optional): Frame to dispatch to.
            toggle (bool, optional): If True then dispatch will be preformed on selection
                and again when deselected. Defaults to False.

        Note:
            Some commands such as ``DefaultNumbering`` require toggling. In such cases
            set arg ``toggle = True``.

            Following Example Sets last three lines to to a numbered list.

            .. code-block:: python

                cursor = Write.get_cursor(doc)
                Write.append_para(cursor, "The following points are important:")
                pos = Write.get_position(cursor)
                Write.append_para(cursor, "Have a good breakfast")
                Write.append_para(cursor, "Have a good lunch")
                Write.append_para(cursor, "Have a good dinner")

                tvc = Write.get_view_cursor(doc)
                tvc.gotoEnd(False)
                Write.dispatch_cmd_left(vcursor=tvc, pos=pos, cmd="DefaultNumbering", toggle=True)

        See Also:
            `LibreOffice Dispatch Commands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_
        """
        curr_pos = mSel.Selection.get_position(vcursor)
        vcursor.goLeft(curr_pos - pos, True)
        mLo.Lo.dispatch_cmd(cmd=cmd, props=props, frame=frame)
        vcursor.goRight(curr_pos - pos, False)
        if toggle is True:
            mLo.Lo.dispatch_cmd(cmd=cmd, props=props, frame=frame)

    # region    style_prev_paragraph()
    @overload
    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor, prop_val: object) -> None:
        """
        Style ParaStyleName of previous paragraph

        Args:
            cursor (XTextCursor): Text Cursor
            prop_val (object): Property value
        """
        ...

    @overload
    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor, prop_val: object, prop_name: str) -> None:
        """
        Style previous paragraph

        Args:
            cursor (XTextCursor): Text Cursor
            prop_val (object): Property value
            prop_name (str): Property Name
        """
        ...

    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor | XParagraphCursor, prop_val: object, prop_name: str = None) -> None:
        """
        Style previous paragraph

        Args:
            cursor (XTextCursor): Text Cursor
            prop_val (object): Property value
            prop_name (str): Property Name

        Example:
            .. code-block:: python

                cursor = Write.get_cursor(doc)
                Write.style_prev_paragraph(cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust")
        """
        if prop_name is None:
            prop_name = "ParaStyleName"
        # raises PropertyNotFoundError if property is not found
        old_val = mProps.Props.get_property(cursor, prop_name)

        cursor.gotoPreviousParagraph(True)  # select previous paragraph
        mProps.Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)

        # reset
        cursor.gotoNextParagraph(False)
        mProps.Props.set_property(prop_set=cursor, name=prop_name, value=old_val)

    # endregion style_prev_paragraph()

    # endregion ---------- text cursor property methods ----------------

    # region ------------- style methods -------------------------------

    @staticmethod
    def get_page_text_width(text_doc: XTextDocument) -> int:
        """
        get the width of the page's text area

        Args:
            text_doc (XTextDocument): Text Document

        Returns:
            int: Page Width on success; Otherwise 0
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            print("Could not access the standard page style")
            return 0

        try:
            width = int(props.getPropertyValue("Width"))
            left_margin = int(props.getPropertyValue("LeftMargin"))
            right_margin = int(props.getPropertyValue("RightMargin"))
            return width - (left_margin + right_margin)
        except Exception as e:
            print("Could not access standard page style dimensions")
            print(f"    {e}")
            return 0

    @staticmethod
    def get_page_size(text_doc: XTextDocument) -> Size:
        """
        Get page size

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            Size: _description_
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")
        try:
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            return Size(width, height)
        except Exception as e:
            raise Exception("Could not access standard page style dimensions") from e

    @staticmethod
    def set_page_format(text_doc: XTextDocument, paper_format: Write.PaperFormat) -> None:
        """
        Set Page Format
        paper_format (:py:attr:`.Write.PaperFormat`): Paper Format.

        Args:
            text_doc (XTextDocument): Text Docuument

        Raises:
            MissingInterfaceError: If text_doc does not implement XPrintable interface

        See Also:
            - :py:meth:`.set_a4_page_format`
        """
        xprintable = mLo.Lo.qi(XPrintable, text_doc, True)
        printer_desc = mProps.Props.make_props(PaperFormat=paper_format)
        xprintable.setPrinter(printer_desc)

    @classmethod
    def set_a4_page_format(cls, text_doc: XTextDocument) -> None:
        """
        Set Page Format to A4

        Args:
            text_doc (XTextDocument): Text Docuument

        See Also:
            :py:meth:`set_page_format`
        """
        cls.set_page_format(text_doc=text_doc, paper_format=Write.PaperFormat.A4)

    # endregion ---------- style methods -------------------------------

    # region ------------- headers and footers -------------------------
    @classmethod
    def set_page_numbers(cls, text_doc: XTextDocument) -> None:
        """
        Modify the footer via the page style for the document.
        Put page number & count in the center of the footer in Times New Roman, 12pt

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PropertiesError: If unable to get properties
            Exception: If Unable to set page numbers
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")

        try:
            props.setPropertyValue("FooterIsOn", True)
            #   footer must be turned on in the document
            footer_text: XText = props.getPropertyValue("FooterText")
            footer_cursor = footer_text.createTextCursor()

            mProps.Props.set_property(
                prop_set=footer_cursor, name="CharFontName", value=mInfo.Info.get_font_general_name()
            )
            mProps.Props.set_property(prop_set=footer_cursor, name="CharHeight", value=12.0)
            mProps.Props.set_property(prop_set=footer_cursor, name="ParaAdjust", value=ParagraphAdjust.CENTER)

            # add text fields to the footer
            pg_number = cls.get_page_number()
            pg_xcontent = mLo.Lo.qi(XTextContent, pg_number)
            if pg_xcontent is None:
                raise mEx.MissingInterfaceError(
                    XTextContent, f"Missing interface for page number. {XTextContent.__pyunointerface__}"
                )
            cls._append_text_content(cursor=footer_cursor, text_content=pg_xcontent)
            cls._append_text(cursor=footer_cursor, text=" of ")
            pg_count = cls.get_page_count()
            pg_count_xcontent = mLo.Lo.qi(XTextContent, pg_number)
            if pg_count_xcontent is None:
                raise mEx.MissingInterfaceError(
                    XTextContent, f"Missing interface for page count. {XTextContent.__pyunointerface__}"
                )
            cls._append_text_content(cursor=footer_cursor, text_content=pg_count_xcontent)
        except Exception as e:
            raise Exception("Unable to set page numbers") from e

    @staticmethod
    def get_page_number() -> XTextField:
        """
        Gets arabic style number showing current page value

        Returns:
            XTextField: Page Number as Text Field
        """
        num_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageNumber")
        mProps.Props.set_property(prop_set=num_field, name="NumberingType", value=NumberingType.ARABIC)
        mProps.Props.set_property(prop_set=num_field, name="SubType", value=PageNumberType.CURRENT)
        return num_field

    @staticmethod
    def get_page_count() -> XTextField:
        """
        return arabic style number showing current page count

        Returns:
            XTextField: Page Count as Text Field
        """
        pc_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageCount")
        mProps.Props.set_property(prop_set=pc_field, name="NumberingType", value=NumberingType.ARABIC)
        return pc_field

    @staticmethod
    def set_header(text_doc: XTextDocument, text: str) -> None:
        """
        Modify the header via the page style for the document.
        Put the text on the right hand side in the header in
        a general font of 10pt.

        Args:
            text_doc (XTextDocument): Text Document
            text (str): Header Text

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to set header text

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_general_name`
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style container")
        try:
            props.setPropertyValue("HeaderIsOn", True)
            # header must be turned on in the document
            # props.setPropertyValue("TopMargin", 2200)
            header_text = mLo.Lo.qi(XText, props.getPropertyValue("HeaderText"))
            header_cursor = header_text.createTextCursor()
            header_cursor.gotoEnd(False)

            header_props = mLo.Lo.qi(XPropertySet, header_cursor, True)
            header_props.setPropertyValue("CharFontName", mInfo.Info.get_font_general_name())
            header_props.setPropertyValue("CharHeight", 10)
            header_props.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT)

            header_text.setString(f"{text}\n")
        except Exception as e:
            raise Exception("Unable to set header text") from e

    @staticmethod
    def get_draw_page(text_doc: XTextDocument) -> XDrawPage:
        """
        Gets draw page

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If text_doc does not implement XDrawPageSupplier interface.

        Returns:
            XDrawPage: Draw Page
        """
        xsupp_page = mLo.Lo.qi(XDrawPageSupplier, text_doc, True)
        return xsupp_page.getDrawPage()

    # endregion ---------- headers and footers -------------------------

    # region ------------- adding elements -----------------------------

    @classmethod
    def add_formula(cls, cursor: XTextCursor, formula: str) -> None:
        """
        Adds a formula

        Args:
            cursor (XTextCursor): Cursor
            formula (str): formula

        Raises:
            CreateInstanceMsfError: If unable to create text.TextEmbeddedObject
            Exception: If unable to add formula
        """
        embed_content = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True)
        try:
            # set class ID for type of object being inserted
            props = mLo.Lo.qi(XPropertySet, embed_content, True)
            props.setPropertyValue("CLSID", mLo.Lo.CLSID.MATH)
            props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER)

            # insert object in document
            cls._append_text_content(cursor=cursor, text_content=embed_content)
            cls.end_line(cursor)

            # access object's model
            embed_obj_supplier = mLo.Lo.qi(XEmbeddedObjectSupplier2, embed_content, True)
            embed_obj_model = embed_obj_supplier.getEmbeddedObject()

            formula_props = mLo.Lo.qi(XPropertySet, embed_obj_model, True)
            formula_props.setPropertyValue("Formula", formula)
            print(f'Inserted formula "{formula}"')
        except Exception as e:
            raise Exception(f'Insertion fo formula "{formula}" failed:') from e

    @classmethod
    def add_hyperlink(cls, cursor: XTextCursor, label: str, url_str: str) -> None:
        """
        Add a hyperlink

        Args:
            cursor (XTextCursor): Text Cursor
            label (str): Hyperlink label
            url_str (str): Hyperlink url

        Raises:
            CreateInstanceMsfError: If unable to create TextField.URL instance
            Exception: If unable to create hyperlink
        """
        try:
            link = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextField.URL")
            if link is None:
                raise ValueError("Null Value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.text.TextField.URL") from e
        try:
            mProps.Props.set_property(prop_set=link, name="URL", value=url_str)
            mProps.Props.set_property(prop_set=link, name="Representation", value=label)

            cls._append_text_content(cursor, link)
            print("Added hyperlink")
        except Exception as e:
            raise Exception("Unable to add hyperlink") from e

    @classmethod
    def add_bookmark(cls, cursor: XTextCursor, name: str) -> None:
        """
        Adds bookmark

        Args:
            cursor (XTextCursor): Text Cursor
            name (str): Bookmark name

        Raises:
            CreateInstanceMsfError: If Unable to create Bookmark instance
            Exception: If unable to add bookmark
        """
        try:
            bmk_content = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.Bookmark")
            if bmk_content is None:
                raise ValueError("Null Value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.text.Bookmark") from e
        try:
            bmk_named = mLo.Lo.qi(XNamed, bmk_content, True)
            bmk_named.setName(name)

            cls._append_text_content(cursor, bmk_content)
        except Exception as e:
            raise Exception("Unable to add bookmark") from e

    @staticmethod
    def find_bookmark(text_doc: XTextDocument, bm_name: str) -> XTextContent | None:
        """
        Finds a bookmark

        Args:
            text_doc (XTextDocument): Text Document
            bm_name (str): Bookmark name

        Raises:
            MissingInterfaceError: if text_doc does not implement XBookmarksSupplier interface

        Returns:
            XTextContent | None: Bookmark if found; Otherwise, None
        """
        supplier = mLo.Lo.qi(XBookmarksSupplier, text_doc, True)

        named_bookmarks = supplier.getBookmarks()
        obookmark = None

        try:
            obookmark = named_bookmarks.getByName(bm_name)
        except Exception:
            print(f"Bookmark '{bm_name}' not found")
            return None
        return mLo.Lo.qi(XTextContent, obookmark)

    @classmethod
    def add_text_frame(
        cls,
        cursor: XTextCursor,
        ypos: int,
        text: str,
        width: int,
        height: int,
        page_num: int = 1,
        border_color: Color | None = CommonColor.RED,
        background_color: Color | None = CommonColor.LIGHT_BLUE,
    ) -> None:
        """
        Adds a text frame with a red border and light blue back color

        Args:
            cursor (XTextCursor): Text Cursor
            ypos (int): Frame Y pos
            text (str): Frame Text
            width (int): Width
            height (int): Height
            page_num (int): Page Number to add text frame
            border_color (Color): Border Color. Defaluts to CommonColor.RED
            background_color (Color): Background Color. Defaluts to CommonColor.LIGHT_BLUE

        Raises:
            CreateInstanceMsfError: If unable to create text.TextFrame
            Exception: If unable to add text frame

        See Also:
            :py:class:`~.utils.color.CommonColor`
        """
        try:
            xframe = mLo.Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame")
            if xframe is None:
                raise ValueError("Null value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextFrame, "com.sun.star.text.TextFrame") from e

        try:
            tf_shape = mLo.Lo.qi(XShape, xframe, True)

            # set dimensions of the text frame
            tf_shape.setSize(Size(width, height))

            #  anchor the text frame
            frame_props = mLo.Lo.qi(XPropertySet, xframe, True)
            frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PAGE)
            frame_props.setPropertyValue("FrameIsAutomaticHeight", True)  # will grow if necessary

            # add a red border around all 4 sides
            border = BorderLine()
            border.OuterLineWidth = 1
            if border_color is not None:
                border.Color = border_color

            frame_props.setPropertyValue("TopBorder", border)
            frame_props.setPropertyValue("BottomBorder", border)
            frame_props.setPropertyValue("LeftBorder", border)
            frame_props.setPropertyValue("RightBorder", border)

            # make the text frame blue
            if background_color is not None:
                frame_props.setPropertyValue("BackTransparent", False)  # not transparent
                frame_props.setPropertyValue("BackColor", background_color)  # light blue

            # Set the horizontal and vertical position
            frame_props.setPropertyValue("HoriOrient", HoriOrientation.RIGHT)
            frame_props.setPropertyValue("VertOrient", VertOrientation.NONE)
            frame_props.setPropertyValue("VertOrientPosition", ypos)  # down from top

            # if page number is Not include for TextContentAnchorType.AT_PAGE
            # then Lo Default so At AT_PARAGRAPH
            frame_props.setPropertyValue("AnchorPageNo", page_num)

            # insert text frame into document (order is important here)
            cls._append_text_content(cursor, xframe)
            cls.end_paragraph(cursor)

            # add text into the text frame
            xframe_text = xframe.getText()
            xtext_range = mLo.Lo.qi(XTextRange, xframe_text.createTextCursor(), True)
            xframe_text.insertString(xtext_range, text, False)
        except Exception as e:
            raise Exception("Insertion of text frame failed:") from e

    @classmethod
    def add_table(
        cls,
        cursor: XTextCursor,
        table_data: Table,
        header_bg_color: Color | None = CommonColor.DARK_BLUE,
        header_fg_color: Color | None = CommonColor.WHITE,
        tbl_bg_color: Color | None = CommonColor.LIGHT_BLUE,
        tbl_fg_color: Color | None = CommonColor.BLACK,
    ) -> None:
        """
        Adds a table.

        Each row becomes a row of the table. The first row is treated as a header.

        Args:
            cursor (XTextCursor): Text Cursor
            table_data (Table): 2D Table with the the first row containing column names.
            header_bg_color (Color | None, optional): Table header background color. Set to None to ignore header color. Defaults to CommonColor.DARK_BLUE.
            header_fg_color (Color | None, optional): Table header forground color. Set to None to ignore header color. Defaults to CommonColor.WHITE.
            tbl_bg_color (Color | None, optional): Table background color. Set to None to ignore background color. Defaults to CommonColor.LIGHT_BLUE.
            tbl_fg_color (Color | None, optional): Table background color. Set to None to ignore background color. Defaults to CommonColor.BLACK.

         Raises:
            ValueError: If table_data is empty
            CreateInstanceMsfError: If unable to create instance of text.TextTable
            Exception: If unable to add table

        See Also:
            :py:class:`~.utils.color.CommonColor`
        """

        def make_cell_name(row: int, col: int) -> str:
            return TableHelper.make_cell_name(row=row + 1, col=col + 1)

        def set_cell_header(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if header_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                mProps.Props.set_property(prop_set=text_cursor, name="CharColor", value=header_fg_color)

            cell_text.setString(str(data))

        def set_cell_text(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if tbl_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                mProps.Props.set_property(prop_set=text_cursor, name="CharColor", value=tbl_fg_color)
            cell_text.setString(str(data))

        num_rows = len(table_data)
        if num_rows == 0:
            raise ValueError("table_data has no values")
        try:
            table = mLo.Lo.create_instance_msf(XTextTable, "com.sun.star.text.TextTable")
            if table is None:
                raise ValueError("Null Value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextTable, "com.sun.star.text.TextTable")

        try:
            num_cols = len(table_data[0])
            print(f"Creating table rows: {num_rows}, cols: {num_cols}")
            table.initialize(num_rows, num_cols)

            # insert the table into the document
            cls._append_text_content(cursor, table)
            cls.end_paragraph(cursor)

            table_props = mLo.Lo.qi(XPropertySet, table, True)

            # set table properties
            if header_bg_color is not None or tbl_bg_color is not None:
                table_props.setPropertyValue("BackTransparent", False)  # not transparent
            if tbl_bg_color is not None:
                table_props.setPropertyValue("BackColor", tbl_bg_color)

            # set color of first row (i.e. the header)
            if header_bg_color is not None:
                rows = table.getRows()
                mProps.Props.set_property(prop_set=rows.getByIndex(0), name="BackColor", value=header_bg_color)

            #  write table header
            row_data = table_data[0]
            for x in range(num_cols):
                set_cell_header(make_cell_name(0, x), row_data[x], table)
                # e.g. "A1", "B1", "C1", etc

            # insert table body
            for y in range(1, num_rows):  # start in 2nd row
                row_data = table_data[y]
                for x in range(num_cols):
                    set_cell_text( make_cell_name(y, x), row_data[x], table)
        except Exception as e:
            raise Exception("Table insertion failed:") from e

    # region    add_image_link()

    @overload
    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr) -> None:
        """
        Add Image Link

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
        """
        ...

    @overload
    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, width: int, height: int) -> None:
        """
        Add Image Link

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
            width (int, optional): Width.
            height (int, optional): Height.
        """
        ...

    @classmethod
    def add_image_link(
        cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, width: int = 0, height: int = 0
    ) -> None:
        """
        Add Image Link

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
            width (int, optional): Width.
            height (int, optional): Height.

        Raises:
            CreateInstanceMsfError: If Unable to create text.TextGraphicObject
            MissingInterfaceError: If unable to obtain XPropertySet interface
            Exception: If unable to add image
        """
        try:
            tgo = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextGraphicObject")
            if tgo is None:
                raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.text.TextGraphicObject")

            props = mLo.Lo.qi(XPropertySet, tgo, True)
            props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER)
            props.setPropertyValue("GraphicURL", mFileIO.FileIO.fnm_to_url(fnm))

            # optionally set the width and height
            if width > 0 and height > 0:
                props.setPropertyValue("Width", width)
                props.setPropertyValue("Height", height)

            # append image to document, followed by a newline
            cls._append_text_content(cursor, tgo)
            cls.end_line(cursor)
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e

    # endregion add_image_link()

    # region    add_image_shape()
    @overload
    @staticmethod
    def add_image_shape(doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr) -> None:
        """
        Add Image Shape

        Currently this method is only suported in terminal. Not in macros.

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
        """
        ...

    @overload
    @staticmethod
    def add_image_shape(doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, width: int, height: int) -> None:
        """
        Add Image Shape

        Currently this method is only suported in terminal. Not in macros.

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
            width (int, optional): Image width
            height (int, optional): Image height
        """
        ...

    @classmethod
    def add_image_shape(
        cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, width: int = 0, height: int = 0
    ) -> None:
        """
        Add Image Shape

        Args:
            doc (XTextDocument): Text Document
            cursor (XTextCursor): Text Cursor
            fnm (PathOrStr): Image path
            width (int, optional): Image width
            height (int, optional): Image height

        Raises:
            CreateInstanceMsfError: If unable to create rawing.GraphicObjectShape
            ValueError: If unable to get image
            MissingInterfaceError: If require interface cannot be obtained.
            Exception: If unable to add image shape
        """

        try:
            if width > 0 and height > 0:
                im_size = Size(width, height)
            else:
                im_size = mImgLo.ImagesLo.get_size_100mm(fnm)  # in 1/100 mm units
                if im_size is None:
                    raise ValueError(f"Unable to get image from {fnm}")

            # create TextContent for an empty graphic
            gos = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.GraphicObjectShape")
            if gos is None:
                raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.drawing.GraphicObjectShape")

            bitmap = mImgLo.ImagesLo.get_bitmap(fnm)
            if bitmap is None:
                raise ValueError(f"Unable to get bitmap of {fnm}")
            # store the image's bitmap as the graphic shape's URL's value
            mProps.Props.set_property(prop_set=gos, name="GraphicURL", value=bitmap)

            # set the shape's size
            xdraw_shape = mLo.Lo.qi(XShape, gos, True)
            xdraw_shape.setSize(im_size)

            # insert image shape into the document, followed by newline
            cls._append_text_content(cursor, gos)
            cls.end_line(cursor)
        except ValueError:
            raise
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e

    # endregion add_image_shape()

    @classmethod
    def add_line_divider(cls, cursor: XTextCursor, line_width: int) -> None:
        """
        Adds a line divider

        Args:
            cursor (XTextCursor): Text Cursor
            line_width (int): Line width

        Raises:
            CreateInstanceMsfError: If unable to create drawing.LineShape instance
            MissingInterfaceError: If unable to obtain XShape interface
            Exception: If unable to add Line divider
        """
        try:
            ls = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.LineShape")
            if ls is None:
                raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.drawing.LineShape")

            line_shape = mLo.Lo.qi(XShape, ls, True)
            line_shape.setSize(Size(line_width, 0))

            cls.end_paragraph(cursor)
            cls._append_text_content(cursor, ls)
            cls.end_paragraph(cursor)

            # center the previous paragraph
            cls.style_prev_paragraph(cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust")

            cls.end_paragraph(cursor)
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception("Insertion of graphic line failed") from e

    # endregion ---------- adding elements -----------------------------

    # region ------------- extracting graphics from text doc -----------

    @classmethod
    def get_text_graphics(cls, text_doc: XTextDocument) -> List[XGraphic]:
        """
        Gets text graphics.

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            Exception: If unable to get text graphics

        Returns:
            List[XGraphic]: Text Graphics

        Note:
            If there is error getting a graphic link then it is ignored
            and not added to the return value.
        """
        try:
            xname_access = cls.get_graphic_links(text_doc)
            if xname_access is None:
                raise ValueError("Unable to get Graphic Links")
            names = xname_access.getElementNames()

            pics: List[XGraphic] = []
            for name in names:
                graphic_link = None
                try:
                    graphic_link = xname_access.getByName(name)
                except UnoException:
                    pass
                if graphic_link is None:
                    print(f"No graphic found for {name}")
                else:
                    try:
                        xgraphic = mImgLo.ImagesLo.load_graphic_link(graphic_link)
                        pics.append(xgraphic)
                    except Exception as e:
                        print(f"{name} could not be accessed:")
                        print(f"    {e}")
            if len(pics) == 0:
                return None
            return pics
        except Exception as e:
            raise Exception(f"Get text graphics failed:") from e

    @staticmethod
    def get_graphic_links(doc: XComponent) -> XNameAccess | None:
        """
        Gets graphic links

        Args:
            doc (XComponent): Document

        Raises:
            MissingInterfaceError: if doc does not implement XTextGraphicObjectsSupplier interface

        Returns:
            XNameAccess | None: Graphic Links on success, Otherwise, None
        """
        ims_supplier = mLo.Lo.qi(XTextGraphicObjectsSupplier, doc, True)

        xname_access = ims_supplier.getGraphicObjects()
        if xname_access is None:
            print("Name access to graphics not possible")
            return None

        if not xname_access.hasElements():
            print("No graphics elements found")
            return None

        return xname_access

    @staticmethod
    def is_anchored_graphic(graphic: object) -> bool:
        """
        Gets if a graphic object is an anchored graphic

        Args:
            graphic (object): object that implements XServiceInfo

        Returns:
            bool: True if is anchored graphic; Otheriwse, False
        """
        serv_info = mLo.Lo.qi(XServiceInfo, graphic)
        return (
            serv_info is not None
            and serv_info.supportsService("com.sun.star.text.TextContent")
            and serv_info.supportsService("com.sun.star.text.TextGraphicObject")
        )

    @staticmethod
    def get_shapes(text_doc: XTextDocument) -> XDrawPage:
        """
        Gets shapes

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If text_doc does not implement XDrawPageSupplier interface

        Returns:
            XDrawPage: shapes
        """
        draw_page_supplier = mLo.Lo.qi(XDrawPageSupplier, text_doc, True)

        return draw_page_supplier.getDrawPage()

    # endregion ---------- extracting graphics from text doc -----------

    # region ------------  Linguistic API ------------------------------

    @classmethod
    def print_services_info(cls, lingo_mgr: XLinguServiceManager2) -> None:
        """
        Prints service info to console

        Args:
            lingo_mgr (XLinguServiceManager2): Serivice manager
        """
        loc = Locale("en", "US", "")
        print("Available Services:")
        cls.print_avail_service_info(lingo_mgr, "SpellChecker", loc)
        cls.print_avail_service_info(lingo_mgr, "Thesaurus", loc)
        cls.print_avail_service_info(lingo_mgr, "Hyphenator", loc)
        cls.print_avail_service_info(lingo_mgr, "Proofreader", loc)

        print("Configured Services:")
        cls.print_config_service_info(lingo_mgr, "SpellChecker", loc)
        cls.print_config_service_info(lingo_mgr, "Thesaurus", loc)
        cls.print_config_service_info(lingo_mgr, "Hyphenator", loc)
        cls.print_config_service_info(lingo_mgr, "Proofreader", loc)

        print()
        cls.print_locales("SpellChecker", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.SpellChecker"))
        cls.print_locales("Thesaurus", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Thesaurus"))
        cls.print_locales("Hyphenator", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Hyphenator"))
        cls.print_locales("Proofreader", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Proofreader"))
        print()

    @staticmethod
    def print_avail_service_info(lingo_mgr: XLinguServiceManager2, service: str, loc: Locale) -> None:
        """
        Prints available service info to console

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manger
            service (str): Service Name
            loc (Locale): Locale
        """
        service_names = lingo_mgr.getAvailableServices(f"com.sun.star.linguistic2.{service}", loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")

    @staticmethod
    def print_config_service_info(lingo_mgr: XLinguServiceManager2, service: str, loc: Locale) -> None:
        """
        Print config service info to console

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager
            service (str): Service Name
            loc (Locale): Locale
        """
        service_names = lingo_mgr.getAvailableServices(f"com.sun.star.linguistic2.{service}", loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")

    @staticmethod
    def print_locales(service: str, loc: Iterable[Locale]) -> None:
        """
        Print locales to console

        Args:
            service (str): Service
            loc (Iterable[Locale]): Locale's
        """
        countries: List[str] = []
        for l in loc:
            countries.append(l.Country)
        countries.sort()

        print(f"Locales for {service} ({len(countries)})")
        for i, country in enumerate(countries):
            if (i % 0) == 0:
                print()
            print(f"  {country}")
        print()
        print()

    @staticmethod
    def set_configured_services(lingo_mgr: XLinguServiceManager2, service: str, impl_name: str) -> None:
        """
        Set configured Services

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager
            service (str): Service Name
            impl_name (str): Service implementation name
        """
        loc = Locale("en", "US", "")
        impl_names = (impl_name,)
        lingo_mgr.setConfiguredServices(f"com.sun.star.linguistic2.{service}", loc, impl_names)

    @classmethod
    def dicts_info(cls) -> None:
        """
        Prints dictionary info to console
        """
        dict_lst = mLo.Lo.create_instance_mcf(XSearchableDictionaryList, "com.sun.star.linguistic2.DictionaryList")
        if not dict_lst:
            print("No list of dictionaries found")
            return
        cls.print_dicts_info(dict_lst)

        cd_list = mLo.Lo.create_instance_mcf(
            XConversionDictionaryList, "com.sun.star.linguistic2.ConversionDictionaryList"
        )
        if cd_list is None:
            print("No list of conversion dictionaries found")
            return
        cls.print_con_dicts_info(cd_list)

    @classmethod
    def print_dicts_info(cls, dict_list: XSearchableDictionaryList) -> None:
        """
        Prints dictionaries info to console

        Args:
            dict_list (XSearchableDictionaryList): dictionary list
        """
        if dict_list is None:
            print("Dictionary list is null")
            return
        print(f"No. of dictionaries: {dict_list.getCount()}")
        dicts = dict_list.getDictionaries()
        for d in dicts:
            print(
                f"  {d.getName()} ({d.getCount()}); ({'active' if d.isActive() else 'na'}); '{d.getLocale().Country}'; {cls.get_dict_type(d.getDictionaryType())}"
            )
        print()

    @staticmethod
    def get_dict_type(dt: DictionaryType) -> str:
        """
        Gets dictionary type

        Args:
            dt (DictionaryType): Dictionary Type

        Returns:
            str: positive, negative, mixed, or ?? if unknown
        """
        if dt == Write.DictionaryType.POSITIVE:
            return "positive"
        if dt == Write.DictionaryType.NEGATIVE:
            return "negative"
        if dt == Write.DictionaryType.MIXED:
            return "mixed"
        return "??"

    @staticmethod
    def print_con_dicts_info(cd_lst: XConversionDictionaryList) -> None:
        """
        Prints Conversion dictionary list to console

        Args:
            cd_lst (XConversionDictionaryList): conversion dictionary list
        """
        if cd_lst is None:
            print("Conversion Dictionary list is null")
            return

        dc_con = cd_lst.getDictionaryContainer()
        dc_names = dc_con.getElementNames()
        print(f"No. of conversion dictionaries: {len(dc_names)}")
        for name in dc_names:
            print(f"  {name}")
        print()

    @staticmethod
    def get_lingu_properties() -> XLinguProperties:
        """
        Gets Lingu Properties

        Raises:
            CreateInstanceMcfError: If unable to create linguistic2.LinguProperties instance

        Returns:
            XLinguProperties: Properties
        """
        props = mLo.Lo.create_instance_mcf(XLinguProperties, "com.sun.star.linguistic2.LinguProperties", raise_err=True)
        return props

    # endregion ---------  Linguistic API ------------------------------

    # region ------------- Linguistics: spell checking -----------------

    @staticmethod
    def load_spell_checker() -> XSpellChecker:
        """
        Gets spell checker

        Raises:
            CreateInstanceMcfError: If unable to create linguistic2.LinguServiceManager instance

        Returns:
            XSpellChecker: spell checker
        """
        lingo_mgr = mLo.Lo.create_instance_mcf(XLinguServiceManager, "com.sun.star.linguistic2.LinguServiceManager", raise_err=True)
        return lingo_mgr.getSpellChecker()

    @classmethod
    def spell_sentence(cls, sent: str, speller: XSpellChecker) -> int:
        """
        Spell Check sentence for en US

        Args:
            sent (str): Setence to spell check
            speller (XSpellChecker): spell checker instance

        Returns:
            int: Number of words spelled incorrectly
        """
        # https://tinyurl.com/y6o8doh2
        words = re.split("\W+", sent)
        count = 0
        for word in words:
            is_correct = cls.spell_word(word, speller)
            count = count + (0 if is_correct else 1)
        return count

    @staticmethod
    def spell_word(word: str, speller: XSpellChecker) -> bool:
        """
        Spell Check a word for en US

        Args:
            word (str): word to spell check
            speller (XSpellChecker): spell checker instance

        Returns:
            bool: True if no spelling errors are detected; Otherwise, False
        """
        loc = Locale("en", "US", "")
        alts = speller.spell(word, loc, tuple())
        if alts is not None:
            print(f"* '{word}' is unknown. Try:")
            alt_words = alts.getAlternatives()
            mLo.Lo.print_names(alt_words)
            return False
        return True

    # endregion ---------- Linguistics: spell checking -----------------

    # region ------------- Linguistics: thesaurus ----------------------

    @staticmethod
    def load_thesaurus() -> XThesaurus:
        """
        Gets Thesaurus

        Raises:
            CreateInstanceMcfError: If unable to create linguistic2.LinguServiceManager instance

        Returns:
            XThesaurus: Thesaurus
        """
        lingo_mgr = mLo.Lo.create_instance_mcf(XLinguServiceManager, "com.sun.star.linguistic2.LinguServiceManager", raise_err=True)
        return lingo_mgr.getThesaurus()

    @staticmethod
    def print_meaning(word: str, thesaurus: XThesaurus) -> int:
        """
        Prints word meanings found in thesaurus to console

        Args:
            word (str): Word to print meanings of
            thesaurus (XThesaurus): thesaurus instance

        Returns:
            int: _description_
        """
        loc = Locale("en", "US", "")
        meanings = thesaurus.queryMeanings(word, loc, tuple())
        if meanings is None:
            print(f"'{word}' NOT found int thesaurus")
            print()
            return 0
        m_len = len(meanings)
        print(f"'{word}' found in thesaurus; number of meanings: {m_len}")

        for i, meaning in enumerate(meanings):
            print(f"{i+1}. Meaning: {meaning.getMeaning()}")
            synonyms = meaning.querySynonyms()
            print(f" No. of  synonyms: {len(synonyms)}")
            for synonym in synonyms:
                print(f"    {synonym}")
            print()
        return m_len

    # endregion ---------- Linguistics: thesaurus ----------------------

    # region ------------- Linguistics: grammar checking ---------------

    @staticmethod
    def load_proofreader() -> XProofreader:
        """
        Gets Proof Reader

        Raises:
            CreateInstanceMcfError: If unable to create linguistic2.Proofreader instance

        Returns:
            XProofreader: Proof Reader
        """
        proof = mLo.Lo.create_instance_mcf(XProofreader, "com.sun.star.linguistic2.Proofreader", raise_err=True)
        return proof

    @classmethod
    def proof_sentence(cls, sent: str, proofreader: XProofreader) -> int:
        """
        Proofs a sentence for en US

        Args:
            sent (str): sentence to proof
            proofreader (XProofreader): Proof reader instance

        Returns:
            int: Number of word of sentence that did not pass proof reading.
        """
        loc = Locale("en", "US", "")
        pr_res = proofreader.doProofreading("1", sent, loc, 0, len(sent), tuple())
        num_errs = 0
        if pr_res is not None:
            errs = pr_res.aErrors
            if len(errs) > 0:
                for err in errs:
                    cls.print_proof_error(err)
                    num_errs += 1
        return num_errs

    @staticmethod
    def print_proof_error(string: str, err: SingleProofreadingError) -> None:
        """
        Prints proof errors to console.

        Args:
            string (str): error string
            err (SingleProofreadingError): Single proof reading error
        """
        e_end = err.nErrorStart + err.nErrorLength
        err_txt = string[err.nErrorStart : e_end]
        print(f"G* {err.aShortComment} in: '{err_txt}'")
        if err.aSuggestions > 0:
            print(f"  Suggested change: '{err.aSuggestions[0]}'")
        print()

    # endregion ---------- Linguistics: grammar checking ---------------

    # region ------------- Linguistics: location guessing --------------

    @staticmethod
    def guess_locale(test_str: str) -> Locale | None:
        """
        Guesses Primary Language and returns results

        Args:
            test_str (str): text used to make guess

        Returns:
            Locale | None: Local if guess succeeds; Otherwise, None
        """
        guesser = mLo.Lo.create_instance_mcf(XLanguageGuessing, "com.sun.star.linguistic2.LanguageGuessing")
        if guesser is None:
            print("No language guesser found")
            return None
        return guesser.guessPrimaryLanguage(test_str, 0, len(test_str))

    @staticmethod
    def print_locale(loc: Locale) -> None:
        """
        Prints a locale to the console

        Args:
            loc (Locale): Locale to print
        """
        if loc is not None:
            print(f"Locale lang: '{loc.Language}'; country: '{loc.Country}'; variant: '{loc.Variant}'")

    # endregion ---------- Linguistics: location guessing --------------

    # region ------------- Linguistics dialogs and menu items ----------

    @staticmethod
    def open_sent_check_options() -> None:
        """open Options - Language Settings - English sentence checking"""
        pip = mInfo.Info.get_pip()
        lang_ext = pip.getPackageLocation("org.openoffice.en.hunspell.dictionaries")
        print(f"Lang Ext: {lang_ext}")
        url = f"{lang_ext}/dialog/en.xdl"
        props = mProps.Props.make_props(OptionsPageURL=url)
        mLo.Lo.dispatch_cmd(cmd="OptionsTreeDialog", props=props)
        mLo.Lo.wait(2000)

    @staticmethod
    def open_spell_grammar_dialog() -> None:
        """activate dialog in  Tools > Speling and Grammar..."""
        mLo.Lo.dispatch_cmd("SpellingAndGrammarDialog")
        mLo.Lo.wait(2000)

    @staticmethod
    def toggle_auto_spell_check() -> None:
        """Toggles spell check on and off"""
        mLo.Lo.dispatch_cmd("SpellOnline")

    @staticmethod
    def open_thesaurus_dialog() -> None:
        """Opens LibreOffice Thesaurus Dialog"""
        mLo.Lo.dispatch_cmd("ThesaurusDialog")

    # endregion ---------- Linguistics dialogs and menu items ----------

    @classmethod
    def print_page_size(cls, text_doc: XTextDocument) -> None:
        """
        Prints Page size to console

        Args:
            text_doc (XTextDocument): Text Document
        """
        # see section 7.17  of Useful Macro Information For OpenOffice By Andrew Pitonyak.pdf
        size = cls.get_page_size(text_doc)
        print("Page Size is:")
        print(f"  {round(size.Width / 100)} mm by {round(size.Height / 100)} mm")
        print(f"  {round(size.Width / 2540)} inches by {round(size.Height / 2540)} inches")
        print(f"  {round((size.Width *72.0) / 2540.0)} picas by {round((size.Height *72.0) / 2540.0)} picas")

    @classmethod
    def _get_left_cursor(cls, o_sel: XTextRange, o_text: XText) -> XTextCursor:
        range_compare = cls.text_range_compare
        if range_compare.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
            o_range = o_sel.getEnd()
        else:
            o_range = o_sel.getStart()
        cursor = o_text.createTextCursorByRange(o_range)
        cursor.goRight(0, False)
        return cursor

    @classmethod
    def _get_right_cursor(cls, o_sel: XTextRange, o_text: XText) -> XTextCursor:

        range_compare = cls.text_range_compare
        if range_compare.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
            o_range = o_sel.getStart()
        else:
            o_range = o_sel.getEnd()
        cursor = o_text.createTextCursorByRange(o_range)
        cursor.goLeft(0, False)
        return cursor


__all__ = ("Write",)
