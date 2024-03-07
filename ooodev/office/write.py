# coding: utf-8
# Python conversion of Write.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING, Iterable, List, Sequence, cast, overload, Union
import re
import uno
from com.sun.star.awt import FontWeight
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XEnumerationAccess
from com.sun.star.container import XNamed
from com.sun.star.document import XDocumentInsertable
from com.sun.star.document import XEmbeddedObjectSupplier2
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.drawing import XShape
from com.sun.star.frame import XModel
from com.sun.star.lang import Locale  # struct class
from com.sun.star.lang import XComponent
from com.sun.star.lang import XServiceInfo
from com.sun.star.linguistic2 import XConversionDictionaryList
from com.sun.star.linguistic2 import XLanguageGuessing
from com.sun.star.linguistic2 import XLinguProperties
from com.sun.star.linguistic2 import XLinguServiceManager
from com.sun.star.linguistic2 import XProofreader
from com.sun.star.linguistic2 import XSearchableDictionaryList
from com.sun.star.linguistic2 import XSpellChecker
from com.sun.star.style import XStyle
from com.sun.star.table import BorderLine  # struct
from com.sun.star.text import HoriOrientation
from com.sun.star.text import VertOrientation
from com.sun.star.text import XBookmarksSupplier
from com.sun.star.text import XPageCursor
from com.sun.star.text import XParagraphCursor
from com.sun.star.text import XText
from com.sun.star.text import XTextContent
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextField
from com.sun.star.text import XTextFrame
from com.sun.star.text import XTextFramesSupplier
from com.sun.star.text import XTextGraphicObjectsSupplier
from com.sun.star.text import XTextRange
from com.sun.star.text import XTextTable
from com.sun.star.text import XTextViewCursor
from com.sun.star.uno import Exception as UnoException
from com.sun.star.util import XCloseable
from com.sun.star.view import XPrintable

from ooo.dyn.style.numbering_type import NumberingTypeEnum as NumberingTypeEnum
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.size import Size as UnoSize  # struct
from ooo.dyn.beans.property_value import PropertyValue
from ooo.dyn.linguistic2.dictionary_type import DictionaryType as DictionaryType
from ooo.dyn.style.break_type import BreakType
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust as ParagraphAdjust
from ooo.dyn.text.control_character import ControlCharacterEnum as ControlCharacterEnum
from ooo.dyn.text.page_number_type import PageNumberType
from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType
from ooo.dyn.view.paper_format import PaperFormat as PaperFormat

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.meta.static_meta import classproperty
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units import UnitT
from ooodev.utils import file_io as mFileIO
from ooodev.utils import gen_util as mUtil
from ooodev.utils import images_lo as mImgLo
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils import selection as mSel
from ooodev.utils.color import CommonColor, Color
from ooodev.utils.data_type.size import Size
from ooodev.utils.table_helper import TableHelper


if TYPE_CHECKING:
    # from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XEnumeration
    from com.sun.star.container import XNameAccess
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XFrame
    from com.sun.star.graphic import XGraphic
    from com.sun.star.linguistic2 import SingleProofreadingError
    from com.sun.star.linguistic2 import XLinguServiceManager2
    from com.sun.star.linguistic2 import XThesaurus
    from com.sun.star.text import XTextCursor
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.type_var import PathOrStr
    from ooodev.utils.type_var import Table
    from ooodev.utils.type_var import DocOrCursor
else:
    StyleT = Any

# endregion Imports


class Write(mSel.Selection):
    # region    Selection Overloads

    # for unknown reason Sphinx docs is not including overloads from inherited class.
    # At least not for static methods. My current work around is to implement the same
    # methods in this class.

    # region    get_cursor()
    # https://tinyurl.com/2dlclzqf
    @overload
    @staticmethod
    def get_cursor() -> XTextCursor:
        """
        Gets text cursor from the current document.

        |lo_unsafe|

        Returns:
            XTextCursor: Cursor
        """
        ...

    @overload
    @staticmethod
    def get_cursor(cursor_obj: DocOrCursor) -> XTextCursor:
        """
        Gets text cursor.

        |lo_safe|

        Args:
            cursor_obj (DocOrCursor): Text Document or Text Cursor

        Returns:
            XTextCursor: Cursor
        """
        ...

    @overload
    @staticmethod
    def get_cursor(rng: XTextRange, txt: XText) -> XTextCursor:
        """
        Gets text cursor.

        |lo_safe|

        Args:
            rng (XTextRange): Text Range Instance
            txt (XText): Text Instance

        Returns:
            XTextCursor: Cursor
        """
        ...

    @overload
    @staticmethod
    def get_cursor(rng: XTextRange, text_doc: XTextDocument) -> XTextCursor:
        """
        Gets text cursor.

        |lo_safe|

        Args:
            rng (XTextRange): Text Range instance
            text_doc (XTextDocument): Text Document instance

        Returns:
            XTextCursor: Cursor
        """
        ...

    @staticmethod
    def get_cursor(*args, **kwargs) -> XTextCursor:
        """
        Gets text cursor.

        Args:
            cursor_obj (DocOrCursor): Text Document or Text View Cursor
            rng (XTextRange): Text Range Instance
            text_doc (XTextDocument): Text Document instance

        Raises:
            CursorError: If Unable to get cursor

        Returns:
            XTextCursor: Cursor

        .. versionchanged:: 0.9.0
            Added overload ``get_cursor()``
        """
        return mSel.Selection.get_cursor(*args, **kwargs)

        # endregion get_cursor()

    # endregion Selection Overloads

    # region ------------- doc / open / close /create/ etc -------------

    # region open_doc()
    @overload
    @classmethod
    def open_doc(cls) -> XTextDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr) -> XTextDocument: ...

    @overload
    @classmethod
    def open_doc(cls, *, loader: XComponentLoader) -> XTextDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XTextDocument: ...

    @classmethod
    def open_doc(cls, fnm: PathOrStr | None = None, loader: XComponentLoader | None = None) -> XTextDocument:
        """
        Opens or creates a Text (Writer) document.

        |lo_unsafe|

        Args:
            fnm (PathOrStr): Writer file to open. If omitted then a new Writer document is returned.
            loader (XComponentLoader): Component loader

        Raises:
            Exception: If Document is Null
            Exception: If Not a Text Document
            MissingInterfaceError: If unable to obtain XTextDocument interface
            CancelEventError: if DOC_OPENING event is canceled.

        Returns:
            XTextDocument: Text Document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``fnm`` and ``loader``.

            If ``fnm`` is omitted then ``DOC_OPENED`` event will not be raised.

        Attention:
            :py:meth:`Lo.open_doc <.utils.lo.Lo.open_doc>` method is called along with any of its events.
        """
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(Write.open_doc.__qualname__)
        cargs.event_data = {"fnm": fnm, "loader": loader}
        _Events().trigger(WriteNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        fnm = cast("PathOrStr", cargs.event_data["fnm"])
        if fnm:
            doc = mLo.Lo.open_doc(fnm=fnm) if loader is None else mLo.Lo.open_doc(fnm=fnm, loader=loader)
        elif loader is None:
            doc = cls.create_doc()
        else:
            doc = cls.create_doc(loader=loader)

        if not cls.is_text(doc):
            mLo.Lo.print(f"Not a text document; closing '{fnm}'")
            mLo.Lo.close_doc(doc)
            raise Exception("Not a text document")
        text_doc = mLo.Lo.qi(XTextDocument, doc)
        if text_doc is None:
            mLo.Lo.print(f"Not a text document; closing '{fnm}'")
            mLo.Lo.close_doc(doc)
            raise mEx.MissingInterfaceError(XTextDocument)
        if fnm:
            _Events().trigger(WriteNamedEvent.DOC_OPENED, EventArgs.from_args(cargs))
        return text_doc

    # endregion open_doc()

    @staticmethod
    def is_text(doc: XComponent) -> bool:
        """
        Gets if doc is an actual Writer document.

        |lo_safe|

        Args:
            doc (XComponent): Document Component

        Returns:
            bool: True if doc is Writer Document; Otherwise, False
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.WRITER)

    # region get_text_doc()

    @overload
    @classmethod
    def get_text_doc(cls) -> XTextDocument:
        """
        Gets a writer document.

        |lo_unsafe|

        When using this method in a macro the ``Lo.get_document()`` value should be passed as ``doc`` arg.

        Args:
            doc (XComponent): Component to get writer document from.

        Returns:
            XTextDocument: Writer document.
        """
        ...

    @overload
    @classmethod
    def get_text_doc(cls, doc: XComponent) -> XTextDocument:
        """
        Gets a writer document.

        |lo_safe|

        When using this method in a macro the ``Lo.get_document()`` value should be passed as ``doc`` arg.

        Args:
            doc (XComponent): Component to get writer document from.

        Returns:
            XTextDocument: Writer document.
        """
        ...

    @classmethod
    def get_text_doc(cls, doc: XComponent | None = None) -> XTextDocument:
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

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_TEXT` :eventref:`src-docs-event`

        .. versionchanged:: 0.9.0
            Added overload ``get_text_doc()``
        """
        if doc is None:
            return mSel.Selection.get_text_doc()
        else:
            return mSel.Selection.get_text_doc(doc=doc)

    # endregion get_text_doc()

    # region create_doc()
    @overload
    @staticmethod
    def create_doc() -> XTextDocument: ...

    @overload
    @staticmethod
    def create_doc(loader: XComponentLoader) -> XTextDocument: ...

    @staticmethod
    def create_doc(loader: XComponentLoader | None = None) -> XTextDocument:
        """
        Creates a new Writer Text Document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader.

        Returns:
            XTextDocument: Text Document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CREATING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CREATED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``loader``.

        Attention:
            :py:meth:`Lo.create_doc <.utils.lo.Lo.create_doc>` method is called along with any of its events.
        """
        cargs = CancelEventArgs(Write.create_doc.__qualname__)
        cargs.event_data = {"loader": loader}
        _Events().trigger(WriteNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if loader:
            doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.WRITER, loader=loader)
        else:
            doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.WRITER)
        _Events().trigger(WriteNamedEvent.DOC_CREATED, EventArgs.from_args(cargs))
        return mLo.Lo.qi(XTextDocument, doc, True)

    # endregion create_doc()

    # region create_doc_from_template()
    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr) -> XTextDocument: ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader) -> XTextDocument: ...

    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader | None = None) -> XTextDocument:
        """
        Create a new Writer Text Document from a template.

        |lo_unsafe|

        Args:
            template_path (PathOrStr): Path to Template.
            loader (XComponentLoader): Component Loader.

        Raises:
            MissingInterfaceError: If Unable to obtain XTextDocument interface.

        Returns:
            XTextDocument: Text Document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CREATING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_TMPL_CREATING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CREATED` :eventref:`src-docs-event`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_TMPL_CREATED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``template_path`` and ``loader``.

        Attention:
            :py:meth:`Lo.create_doc_from_template <.utils.lo.Lo.create_doc_from_template>` method is called along with any of its events.
        """
        cargs = CancelEventArgs(Write.create_doc_from_template.__qualname__)
        cargs.event_data = {"template_path": template_path, "loader": loader}
        _Events().trigger(WriteNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        template_path = cast("PathOrStr", cargs.event_data["template_path"])
        _Events().trigger(WriteNamedEvent.DOC_TMPL_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        if loader:
            doc = mLo.Lo.create_doc_from_template(template_path=template_path, loader=loader)
        else:
            doc = mLo.Lo.create_doc_from_template(template_path=template_path)

        eargs = EventArgs.from_args(cargs)
        _Events().trigger(WriteNamedEvent.DOC_CREATED, eargs)
        _Events().trigger(WriteNamedEvent.DOC_TMPL_CREATED, eargs)
        return mLo.Lo.qi(XTextDocument, doc, True)

    # endregion create_doc_from_template()

    # region close_doc()
    @overload
    @classmethod
    def close_doc(cls) -> bool: ...

    @overload
    @classmethod
    def close_doc(cls, text_doc: XTextDocument) -> bool: ...

    @classmethod
    def close_doc(cls, text_doc: XTextDocument | None = None) -> bool:
        """
        Closes text document.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            MissingInterfaceError: If unable to obtain XCloseable from text_doc.

        Returns:
            bool: False if DOC_CLOSING event is canceled, Other.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc``.

        Attention:
            :py:meth:`Lo.close <.utils.lo.Lo.close>` method is called along with any of its events.

        .. versionchanged:: 0.9.0
            Added overload ``close_doc()``
        """
        if text_doc is None:
            text_doc = cls.active_doc
        cargs = CancelEventArgs(Write.close_doc.__qualname__)
        cargs.event_data = {"text_doc": text_doc}
        _Events().trigger(WriteNamedEvent.DOC_CLOSING, cargs)
        if cargs.cancel:
            return False
        closable = mLo.Lo.qi(XCloseable, cargs.event_data["text_doc"], True)
        result = mLo.Lo.close(closable)
        _Events().trigger(WriteNamedEvent.DOC_CLOSED, EventArgs.from_args(cargs))
        return result

    # endregion close_doc()

    @staticmethod
    def save_doc(text_doc: XTextDocument, fnm: PathOrStr) -> bool:
        """
        Saves text document.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document
            fnm (PathOrStr): Path to save as

        Raises:
            MissingInterfaceError: If text_doc does not implement XComponent interface

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
        cargs = CancelEventArgs(Write.save_doc.__qualname__)
        cargs.event_data = {"text_doc": text_doc, "fnm": fnm}
        _Events().trigger(WriteNamedEvent.DOC_SAVING, cargs)

        if cargs.cancel:
            return False
        fnm = cast("PathOrStr", cargs.event_data["fnm"])

        doc = mLo.Lo.qi(XComponent, text_doc, True)
        result = mLo.Lo.save_doc(doc=doc, fnm=fnm)

        _Events().trigger(WriteNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return result

    # region open_flat_doc_using_text_template()

    @overload
    @classmethod
    def open_flat_doc_using_text_template(cls, fnm: PathOrStr, template_path: PathOrStr) -> XTextDocument: ...

    @overload
    @classmethod
    def open_flat_doc_using_text_template(
        cls, fnm: PathOrStr, template_path: PathOrStr, loader: XComponentLoader
    ) -> XTextDocument: ...

    @classmethod
    def open_flat_doc_using_text_template(
        cls, fnm: PathOrStr, template_path: PathOrStr, loader: XComponentLoader | None = None
    ) -> XTextDocument:
        """
        Open a new text document applying the template as formatting to the flat XML file.

        |lo_unsafe|

        Args:
            fnm (PathOrStr): path to file.
            template_path (PathOrStr): Path to template file (ott).
            loader (XComponentLoader): Component Loader.

        Raises:
            UnOpenableError: If fnm is not able to be opened.
            ValueError: If template_path is not ott file.
            MissingInterfaceError: If template_path document does not implement ``XTextDocument`` interface.
            ValueError: If unable to obtain cursor object.
            Exception: Any other errors

        Returns:
            XTextDocument | None: Text Document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``fnm``, ``template_path`` and ``loader``.

        Attention:
            :py:meth:`Lo.create_doc_from_template <.utils.lo.Lo.create_doc_from_template>` method is called along with any of its events.
        """
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(Write.open_flat_doc_using_text_template.__qualname__)
        cargs.event_data = {"fnm": fnm, "template_path": template_path, "loader": loader}
        _Events().trigger(WriteNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        fnm = cast("PathOrStr", cargs.event_data["fnm"])
        template_path = cast("PathOrStr", cargs.event_data["template_path"])
        if fnm is None:
            raise ValueError("Filename is null")
        pth = mFileIO.FileIO.get_absolute_path(fnm)

        open_file_url = None
        if mFileIO.FileIO.is_openable(pth):
            open_file_url = mFileIO.FileIO.fnm_to_url(pth)

        elif mLo.Lo.is_url(pth):
            mLo.Lo.print(f"Treating filename as a URL: '{pth}'")
            open_file_url = str(pth)
        else:
            raise mEx.UnOpenableError(pth)
        template_ext = mInfo.Info.get_ext(template_path)
        if template_ext != "ott":
            raise ValueError(f"Can only apply a text template as formatting. Not an ott file: {template_path}")

        if loader is None:
            doc = mLo.Lo.create_doc_from_template(template_path=template_path)
        else:
            doc = mLo.Lo.create_doc_from_template(template_path=template_path, loader=loader)
        text_doc = mLo.Lo.qi(XTextDocument, doc)
        if text_doc is None:
            raise mEx.MissingInterfaceError(
                XTextDocument, f"Template is not a text document. Missing: {XTextDocument.__pyunointerface__}"
            )

        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            raise ValueError(f"Unable to get cursor: '{pth}'")

        try:
            cursor.gotoEnd(True)
            di = mLo.Lo.qi(XDocumentInsertable, cursor, True)
            # XDocumentInsertable only works with text files
            di.insertDocumentFromURL(open_file_url, ())
            # Props.makeProps("FilterName", "OpenDocument Text Flat XML"))
            # these props do not work
        except Exception as e:
            raise Exception("Could not insert document") from e
        _Events().trigger(WriteNamedEvent.DOC_OPENED, EventArgs.from_args(cargs))
        return text_doc

    # endregion open_flat_doc_using_text_template()

    @staticmethod
    def get_doc_settings() -> XPropertySet:
        """
        Gets Text Document Settings.

        |lo_unsafe|

        Returns:
            XPropertySet: Settings

        See Also:
            `API DocumentSettings Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1DocumentSettings.html>`__

        .. versionadded:: 0.9.7
        """
        return mLo.Lo.create_instance_msf(XPropertySet, "com.sun.star.text.DocumentSettings", raise_err=True)

    # endregion ---------- doc / open / close /create/ etc -------------

    # region ------------- page methods --------------------------------
    @classmethod
    def get_page_cursor(cls, view_cursor_obj: XTextDocument | XTextViewCursor) -> XPageCursor:
        """
        Get Page cursor.

        |lo_safe|

        Makes it possible to perform cursor movements between pages.

        Args:
            text_doc (XTextDocument | XTextViewCursor): Text Document or View Cursor.

        Raises:
            PageCursorError: If Unable to get cursor.

        Returns:
            XPageCursor: Page Cursor.

        See Also:
            `LibreOffice API XPageCursor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XPageCursor.html>`_
        """
        try:
            view_cursor = mLo.Lo.qi(XTextViewCursor, view_cursor_obj)
            if view_cursor is None:
                view_cursor = cls.get_view_cursor(view_cursor_obj)  # type: ignore
            return mLo.Lo.qi(XPageCursor, view_cursor, True)
        except Exception as e:
            raise mEx.PageCursorError(str(e)) from e

    @staticmethod
    def get_current_page(tv_cursor: XTextViewCursor) -> int:
        """
        Gets the current page.

        |lo_safe|

        Args:
            tv_cursor (XTextViewCursor): Text view Cursor

        Returns:
            int: Page number if present; Otherwise, -1

        See Also:
            :py:meth:`~.Write.get_page_number`
        """
        page_cursor = mLo.Lo.qi(XPageCursor, tv_cursor)
        if page_cursor is None:
            mLo.Lo.print("Could not create a page cursor")
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

        |lo_safe|

        Args:
            tv_cursor (XTextViewCursor): Text View Cursor

        Returns:
            str: coordinates as string

        Warning:
            The ``X`` is relative to the document window and not the document page.
            This means when the document window size changes the ``X`` coordinate will change even if the cursor has not moved.
            This is also the case if the document zoom changes.

            When the document page is zoomed all the way to fill the document window the ``X`` coordinate
            is ``0`` when at the left hand page margin (beginning of a line).

            The ``Y`` coordinate is relative to the top of the document window and not the top of the document page.
            The ``Y`` coordinate seems not to be affected by Document Zoom or scroll position.
        """
        pos = tv_cursor.getPosition()
        return f"({pos.X}, {pos.Y})"

    @staticmethod
    def get_num_of_pages(text_doc: XTextDocument) -> int:
        """
        Gets document page count.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            MissingInterfaceError: If text_doc does not implement XModel interface.

        Returns:
            int: page count.
        """
        model = mLo.Lo.qi(XModel, text_doc, True)
        controller = model.getCurrentController()
        return int(mProps.Props.get(controller, "PageCount"))

    @classmethod
    def print_page_size(cls, text_doc: XTextDocument) -> None:
        """
        Prints Page size to console.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document
        """
        cargs = CancelEventArgs(Write.print_page_size.__qualname__)
        _Events().trigger(GblNamedEvent.PRINTING, cargs)
        if cargs.cancel:
            return
        # see section 7.17  of Useful Macro Information For OpenOffice By Andrew Pitonyak.pdf
        size = cls.get_page_size(text_doc)
        print("Page Size is:")
        print(f"  {round(size.width / 100)} mm by {round(size.height / 100)} mm")
        print(f"  {round(size.width / 2540)} inches by {round(size.height / 2540)} inches")
        print(f"  {round((size.width *72.0) / 2540.0)} picas by {round((size.height *72.0) / 2540.0)} picas")

    # endregion ---------- page methods --------------------------------

    # region ------------- text writing methods ------------------------

    # region    append()
    @classmethod
    def _append_text(cls, cursor: XTextCursor, text: str) -> None:
        """LO Safe Method."""
        cursor.setString(text)
        cursor.gotoEnd(False)

    @classmethod
    def _append_text_style(cls, cursor: XTextCursor, text: str, styles: Sequence[StyleT]) -> None:
        """LO Safe Method."""
        s_len = len(text)
        if s_len == 0:
            return

        cursor.setString(text)
        cursor.gotoEnd(False)
        style_srv = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.drawing.FillProperties",
        )
        for style in styles:
            if not style.support_service(*style_srv):
                mLo.Lo.print(f"_append_text_style(), Supported services are {style_srv}. Not Supported style: {style}")
                continue
            cargs = CancelEventArgs("Write.append")
            cargs.event_data = style
            _Events().trigger(WriteNamedEvent.STYLING, cargs)
            if cargs.cancel:
                continue
            bak = not any((FormatKind.PARA in style.prop_format_kind, FormatKind.STATIC in style.prop_format_kind))

            if bak:
                # store properties about to be changed
                style.backup(cursor)
            cursor.goLeft(s_len, True)

            style.apply(cursor)

            cursor.gotoEnd(False)
            if bak and style.prop_has_backup:
                try:
                    style.restore(cursor, True)
                except mEx.MultiError as e:
                    mLo.Lo.print("Write.append(): Unable to restore Property")
                    for err in e.errors:
                        mLo.Lo.print(f"  {err}")
                except Exception as e:
                    mLo.Lo.print("Write.append(): Unable to restore Property")
                    mLo.Lo.print(f"  {e}")

            _Events().trigger(WriteNamedEvent.STYLED, EventArgs.from_args(cargs))

    @classmethod
    def _append_ctl_char(cls, cursor: XTextCursor, ctl_char: int) -> None:
        """LO Safe Method"""
        xtext = cursor.getText()
        xtext.insertControlCharacter(cursor, ctl_char, False)
        cursor.gotoEnd(False)

    @classmethod
    def _append_text_content(cls, cursor: XTextCursor, text_content: XTextContent) -> None:
        """Lo Safe Method."""
        xtext = cursor.getText()
        xtext.insertTextContent(cursor, text_content, False)
        cursor.gotoEnd(False)

    @overload
    @classmethod
    def append(cls, cursor: XTextCursor, text: str) -> None: ...

    @overload
    @classmethod
    def append(cls, cursor: XTextCursor, text: str, styles: Sequence[StyleT]) -> None:
        """
        Append content to cursor.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str): Text to append.
            styles (Sequence[StyleT]):One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append(cls, cursor: XTextCursor, ctl_char: ControlCharacterEnum) -> None:
        """
        Append content to cursor.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            ctl_char (int): Control Char (like a paragraph break or a hard space).

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append(cls, cursor: XTextCursor, text_content: XTextContent) -> None:
        """
        Append content to cursor.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text_content (XTextContent): Text content, such as a text table, text frame or text field.

        Returns:
            None:
        """
        ...

    @classmethod
    def append(cls, *args, **kwargs) -> None:
        """
        Append content to cursor.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str): Text to append.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            ctl_char (int): Control Char (like a paragraph break or a hard space).
            text_content (XTextContent): Text content, such as a text table, text frame or text field.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.char </src/format/ooodev.format.writer.direct.char>`

        See Also:
            `API ControlCharacter <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1ControlCharacter.html>`_

        .. versionchanged:: 0.9.0
            Added ``append(cursor: XTextCursor, text: str, styles: Sequence[StyleT])`` overload.

            Added Events.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cursor", "text", "ctl_char", "text_content", "styles")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("append() got an unexpected keyword argument")
            ka[1] = kwargs.get("cursor", None)
            keys = ("text", "ctl_char", "text_content")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("styles", None)
            return ka

        if count not in (2, 3):
            raise TypeError("append() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        cursor = cast("XTextCursor", kargs[1])

        if count == 3:
            cls._append_text_style(cursor=cursor, text=kargs[2], styles=kargs[3])
            return

        if isinstance(kargs[2], str):
            cls._append_text(cursor=cursor, text=kargs[2])
        elif isinstance(kargs[2], int):
            cls._append_ctl_char(cursor=cursor, ctl_char=kargs[2])
        else:
            cls._append_text_content(cursor=cursor, text_content=kargs[2])

    # endregion append()

    # region append_line()
    @overload
    @classmethod
    def append_line(cls, cursor: XTextCursor) -> None:
        """
        Appends a new Line.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append_line(cls, cursor: XTextCursor, text: str) -> None:
        """
        Appends a new Line.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): text to append before new line is inserted.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append_line(cls, cursor: XTextCursor, text: str, styles: Sequence[StyleT]) -> None:
        """
        Appends a new Line.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): text to append before new line is inserted.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is omitted then this argument is ignored.

        Returns:
            None:
        """
        ...

    @classmethod
    def append_line(cls, cursor: XTextCursor, text: str = "", styles: Sequence[StyleT] | None = None) -> None:
        """
        Appends a new Line.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): text to append before new line is inserted.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is omitted then this argument is ignored.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

        .. versionchanged:: 0.9.0
            Added overload ``append_line(cursor: XTextCursor, text: str, styles: Sequence[StyleT])``.

            Added events.
        """
        if text:
            if styles is None:
                cls._append_text(cursor=cursor, text=text)
            else:
                cls._append_text_style(cursor=cursor, text=text, styles=styles)

        cls._append_ctl_char(cursor=cursor, ctl_char=ControlCharacterEnum.LINE_BREAK)

    # endregion append_line()

    @classmethod
    def append_date_time(cls, cursor: XTextCursor) -> None:
        """
        Append two DateTime fields, one for the date, one for the time.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
        """
        dt_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
        mProps.Props.set(dt_field, IsDate=True)  # so date is reported
        xtext_content = mLo.Lo.qi(XTextContent, dt_field, True)
        cls._append_text_content(cursor, xtext_content)
        cls.append(cursor, "; ")

        dt_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.DateTime")
        mProps.Props.set(dt_field, IsDate=False)  # so time is reported
        xtext_content = mLo.Lo.qi(XTextContent, dt_field, True)
        cls._append_text_content(cursor, xtext_content)

    # region append_para()
    @overload
    @classmethod
    def append_para(cls, cursor: XTextCursor) -> None:
        """
        Appends text (if present) and then a paragraph break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append_para(cls, cursor: XTextCursor, text: str) -> None:
        """
        Appends text (if present) and then a paragraph break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): Text to append.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def append_para(cls, cursor: XTextCursor, text: str, styles: Sequence[StyleT]) -> None:
        """
        Appends text (if present) and then a paragraph break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): Text to append.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is empty then this argument is ignored.

        Returns:
            None:
        """
        ...

    @classmethod
    def append_para(cls, cursor: XTextCursor, text: str = "", styles: Sequence[StyleT] | None = None) -> None:
        """
        Appends text (if present) and then a paragraph break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): Text to append.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is empty then this argument is ignored.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.char </src/format/ooodev.format.writer.direct.char>`
            - :doc:`ooodev.format.writer.direct.para </src/format/ooodev.format.writer.direct.para>`

        .. versionchanged:: 0.9.0
            Added overload ``append_para(cursor: XTextCursor, text: str, styles: Sequence[StyleT])``.

            Added Events.
        """

        # paragraph styles need to capture the current paragraph setting and restore them.
        # other styles are handled by _append_text_style().

        # paragraph break (UNICODE 0x000D). \r
        # https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Control_Characters

        restore = False

        style_lst: List[StyleT] = []
        fill_lst: List[StyleT] = []
        restore_style_lst: List[StyleT] = []
        restore_fill_lst: List[StyleT] = []
        style_srv = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.drawing.FillProperties",
        )

        if styles:
            for style in styles:
                if not style.support_service(*style_srv):
                    mLo.Lo.print(f"append_para(), Supported services are {style_srv}. Not Supported style: {style}")
                    continue
                if FormatKind.TXT_CONTENT in style.prop_format_kind and FormatKind.PARA in style.prop_format_kind:
                    fill_lst.append(style)
                else:
                    style_lst.append(style)

        para_c = None
        if fill_lst:
            para_c = cls.get_paragraph_cursor(cursor)

        def capture_old_val(sty: StyleT) -> None:
            nonlocal restore, para_c, restore_style_lst, restore_fill_lst
            if FormatKind.PARA in sty.prop_format_kind and FormatKind.STATIC not in sty.prop_format_kind:
                restore = True
                if para_c is not None and FormatKind.TXT_CONTENT in sty.prop_format_kind:
                    sty.backup(para_c.TextParagraph)  # type: ignore
                    restore_fill_lst.append(sty)
                else:
                    sty.backup(cursor)
                    restore_style_lst.append(sty)

        if text:
            if styles is None:
                cls._append_text(cursor=cursor, text=text)
            else:
                for style in styles:
                    capture_old_val(style)
                cls._append_text_style(cursor=cursor, text=text, styles=style_lst)

        if fill_lst and para_c is not None:
            # para_c = cls.get_paragraph_cursor(cursor)
            para_c.gotoStartOfParagraph(False)
            para_c.gotoEndOfParagraph(True)
            for style in fill_lst:
                style.apply(para_c.TextParagraph)  # type: ignore
            para_c.gotoEnd(False)

        cls._append_ctl_char(cursor=cursor, ctl_char=ControlCharacterEnum.PARAGRAPH_BREAK)

        for style in restore_style_lst:
            try:
                style.restore(cursor, True)
            except mEx.MultiError as e:
                mLo.Lo.print("Write.append_para(): Unable to restore Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
            except Exception as e:
                mLo.Lo.print("Write.append_para(): Unable to restore Property")
                mLo.Lo.print(f"  {e}")

        if para_c is not None:
            for style in restore_fill_lst:
                try:
                    style.restore(para_c.TextParagraph, True)  # type: ignore
                except mEx.MultiError as e:
                    mLo.Lo.print("Write.append_para(): Unable to restore Property")
                    for err in e.errors:
                        mLo.Lo.print(f"  {err}")
                except Exception as e:
                    mLo.Lo.print("Write.append_para(): Unable to restore Property")
                    mLo.Lo.print(f"  {e}")

    # endregion append_para()

    @classmethod
    def end_line(cls, cursor: XTextCursor) -> None:
        """
        Inserts a line break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor
        """
        cls._append_ctl_char(cursor=cursor, ctl_char=ControlCharacterEnum.LINE_BREAK)

    @classmethod
    def end_paragraph(cls, cursor: XTextCursor) -> None:
        """
        Inserts a paragraph break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor
        """
        cls._append_ctl_char(cursor=cursor, ctl_char=ControlCharacterEnum.PARAGRAPH_BREAK)

    @classmethod
    def page_break(cls, cursor: XTextCursor) -> None:
        """
        Inserts a page break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor
        """
        mProps.Props.set(cursor, BreakType=BreakType.PAGE_AFTER)
        cls.end_paragraph(cursor)

    @classmethod
    def column_break(cls, cursor: XTextCursor) -> None:
        """
        Inserts a column break.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor
        """
        mProps.Props.set(cursor, BreakType=BreakType.COLUMN_AFTER)
        cls.end_paragraph(cursor)

    @classmethod
    def insert_para(cls, cursor: XTextCursor, para: str, para_style: str) -> None:
        """
        Inserts a paragraph with a style applied.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor
            para (str): Paragraph text
            para_style (str): Style such as 'Heading 1'
        """
        xtext = cursor.getText()
        xtext.insertString(cursor, para, False)
        xtext.insertControlCharacter(cursor, ControlCharacterEnum.PARAGRAPH_BREAK, False)
        cls.style_prev_paragraph(cursor, para_style)

    # endregion ---------- text writing methods ------------------------

    # region ------------- extract text from document ------------------

    @staticmethod
    def split_paragraph_into_sentences(paragraph: str) -> List[str]:
        """
        Alternative method for breaking a paragraph into sentences and return a list.

        ``XSentenceCursor`` occasionally does not divide a paragraph into the correct number of sentences; sometimes two sentences were treated as one.

        |lo_safe|

        Args:
            paragraph (str): input string.

        Returns:
            List[str]: List of string.

        See Also:
            `split paragraph into sentences with regular expressions <https://pythonicprose.blogspot.com/2009/09/python-split-paragraph-into-sentences.html>`_
        """

        # https://pythonicprose.blogspot.com/2009/09/python-split-paragraph-into-sentences.html
        # abbreviations are fairly common.
        # To try and handle that scenario I would change my regular expression to
        # read: '[.!?][\s]{1,2}(?=[A-Z])'
        #   regular expressions are easiest (and fastest)
        sentence_enders = re.compile(r"[.!?][\s]{1,2}")
        return sentence_enders.split(paragraph)

    @staticmethod
    def get_all_text(cursor: XTextCursor) -> str:
        """
        Gets the text part of the document.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.

        Returns:
            str: text.
        """
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        text = cursor.getString()
        cursor.gotoEnd(False)  # to deselect everything
        return text

    @staticmethod
    def get_enumeration(obj: Any) -> XEnumeration:
        """
        Gets Enumeration access from obj.

        |lo_safe|

        Used to enumerate objects in a container which contains objects.

        Args:
            obj (Any): object that implements XEnumerationAccess or XTextDocument.

        Raises:
            MissingInterfaceError: if obj does not implement XEnumerationAccess interface.

        Returns:
            XEnumeration: Enumerator.
        """
        enum_access = mLo.Lo.qi(XEnumerationAccess, obj)
        if enum_access is None:
            # try for XTextDocument
            with contextlib.suppress(AttributeError):
                xtext = obj.getText()  # type: ignore
                if xtext is not None:
                    enum_access = mLo.Lo.qi(XEnumerationAccess, xtext)
        if enum_access is None:
            raise mEx.MissingInterfaceError(XEnumerationAccess)
        return enum_access.createEnumeration()

    # endregion ---------- extract text from document ------------------

    # region ------------- text cursor property methods ----------------

    @classmethod
    def style_left_bold(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles bold from current cursor position left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Number of positions to go left.
        """
        cls.style_left(cursor, pos, "CharWeight", FontWeight.BOLD)

    @classmethod
    def style_left_italic(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles italic from current cursor position left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Number of positions to go left.
        """
        cls.style_left(cursor, pos, "CharPosture", FontSlant.ITALIC)

    @classmethod
    def style_left_color(cls, cursor: XTextCursor, pos: int, color: Color) -> None:
        """
        Styles color from current cursor position left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Number of positions to go left.
            color (~ooodev.utils.color.Color): RGB color as int to apply.

        Returns:
            None:

        See Also:
            :py:class:`~.utils.color.CommonColor`
        """
        cls.style_left(cursor, pos, "CharColor", color)

    @classmethod
    def style_left_code(cls, cursor: XTextCursor, pos: int) -> None:
        """
        Styles using a Mono font from current cursor position left by pos amount.
        Font Char Height is set to ``10``

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Number of positions to go left.

        Returns:
            None:

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_mono_name`
        """
        cls.style_left(cursor, pos, "CharFontName", mInfo.Info.get_font_mono_name())
        cls.style_left(cursor, pos, "CharHeight", 10)

    # region style()
    @classmethod
    def _style(
        cls, pos: int, distance: int, prop_name: str, prop_val: object, cursor: XTextCursor | None = None
    ) -> None:
        """If Cursor then Lo Safe Method, otherwise Lo Unsafe Method."""
        cargs = KeyValCancelArgs("Write.style", prop_name, prop_val)
        _Events().trigger(WriteNamedEvent.STYLING, cargs)
        if cargs.cancel:
            return
        if cursor is None:
            cursor = cls.get_cursor()
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(distance, True)
        mProps.Props.set(cursor, **{prop_name: prop_val})
        cursor.gotoEnd(False)
        _Events().trigger(WriteNamedEvent.STYLED, KeyValArgs.from_args(cargs))  # type: ignore

    @classmethod
    def _style_style(
        cls, pos: int, distance: int, styles: Sequence[StyleT], cursor: XTextCursor | None = None
    ) -> None:
        """If Cursor then Lo Safe Method, otherwise Lo Unsafe Method."""
        if cursor is None:
            cursor = cls.get_cursor()
        # cursor.collapseToEnd()
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(distance, True)

        style_srv = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.drawing.FillProperties",
        )

        for style in styles:
            if not style.support_service(*style_srv):
                mLo.Lo.print(f"_style_style(), Supported services are {style_srv}. Not Supported style: {style}")
                continue
            cargs = CancelEventArgs("Write.style")
            cargs.event_data = style
            _Events().trigger(WriteNamedEvent.STYLING, cargs)
            if cargs.cancel:
                continue
            style.apply(cursor)
            _Events().trigger(WriteNamedEvent.STYLED, EventArgs.from_args(cargs))
        cursor.gotoEnd(False)

    @overload
    @classmethod
    def style(cls, pos: int, length: int, styles: Sequence[StyleT]) -> None:
        """
        Styles. From position styles right by distance amount.

        |lo_unsafe|

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style(cls, pos: int, length: int, styles: Sequence[StyleT], cursor: XTextCursor) -> None:
        """
        Styles. From position styles right by distance amount.

        |lo_safe|

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            cursor (XTextCursor): Text Cursor

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style(cls, pos: int, length: int, prop_name: str, prop_val: object) -> None:
        """
        Styles. From position styles right by distance amount.

        |lo_unsafe|

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style(cls, pos: int, length: int, prop_name: str, prop_val: object, cursor: XTextCursor) -> None:
        """
        Styles. From position styles right by distance amount.

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``
            cursor (XTextCursor): Text Cursor

        Returns:
            None:
        """
        ...

    @classmethod
    def style(cls, *args, **kwargs) -> None:
        """
        Styles. From position styles right by distance amount.

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``
            cursor (XTextCursor): Text Cursor

        Returns:
            None:

        See Also:
            :py:meth:`~.Write.style_left`

        Note:
            Unlike :py:meth:`~.Write.style_left` this method does not restore any style properties after style is applied.

        .. versionadded:: 0.9.0
        """
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("pos", "length", "prop_name", "prop_val", "styles", "cursor")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("style() got an unexpected keyword argument")
            ka[1] = kwargs.get("pos", None)
            ka[2] = kwargs.get("length", None)
            keys = ("prop_name", "styles")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            keys = ("prop_val", "cursor")
            for key in keys:
                if key in kwargs:
                    ka[4] = kwargs[key]
                    break
            if count == 4:
                return ka
            ka[5] = kwargs.get("cursor", None)
            return ka

        if count not in (3, 4, 5):
            raise TypeError("style() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if count == 3:
            return cls._style_style(pos=kargs[1], distance=kargs[2], styles=kargs[3])
        if count == 4:
            # cursor or prop value
            arg4 = kargs[4]
            if mInfo.Info.is_uno(arg4):
                return cls._style_style(pos=kargs[1], distance=kargs[2], styles=kargs[3], cursor=arg4)
            return cls._style(pos=kargs[1], distance=kargs[2], prop_name=kargs[3], prop_val=arg4)
        return cls._style(pos=kargs[1], distance=kargs[2], prop_name=kargs[3], prop_val=kargs[4], cursor=kargs[5])

    # endregion style()
    # region style_left()

    @classmethod
    def _style_left(cls, cursor: XTextCursor, pos: int, prop_name: str, prop_val: object) -> None:
        """LO UN-safe Method"""
        cargs = KeyValCancelArgs("Write.style", prop_name, prop_val)
        cargs.event_data = {"pos": pos}
        _Events().trigger(WriteNamedEvent.STYLING, cargs)
        if cargs.cancel:
            return
        old_val = None
        if pos == 0:
            cursor.goLeft(0, True)
            amt = 0
        else:
            old_val = mProps.Props.get(cursor, prop_name)
            curr_pos = mSel.Selection.get_position(cursor)
            amt = curr_pos - pos
            cursor.goLeft(amt, True)
        mProps.Props.set(cursor, **{prop_name: prop_val})

        if old_val is not None and pos > 0:
            cursor.goRight(amt, False)
            mProps.Props.set(cursor, **{prop_name: old_val})
        else:
            cursor.goRight(0, False)
        _Events().trigger(WriteNamedEvent.STYLED, KeyValArgs.from_args(cargs))  # type: ignore

    @classmethod
    def _style_left_style(cls, cursor: XTextCursor, pos: int, styles: Sequence[StyleT]) -> None:
        """LO UN-safe Method"""
        # store properties about to be changed

        if pos == 0:
            cursor.goLeft(0, True)
            amt = 0
        else:
            curr_pos = mSel.Selection.get_position(cursor)
            amt = curr_pos - pos

        style_srv = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.drawing.FillProperties",
        )

        for style in styles:
            if not style.support_service(*style_srv):
                mLo.Lo.print(f"_style_left_style(), Supported services are {style_srv}. Not Supported style: {style}")
                continue
            cargs = CancelEventArgs("Write.style_left")
            cargs.event_data = style
            _Events().trigger(WriteNamedEvent.STYLING, cargs)
            if cargs.cancel:
                continue
            bak = None
            if pos > 0:
                bak = not any((FormatKind.PARA in style.prop_format_kind, FormatKind.STATIC in style.prop_format_kind))
                if bak:
                    style.backup(cursor)

                cursor.goLeft(amt, True)
            style.apply(cursor)
            if pos > 0:
                cursor.goRight(amt, False)

            if bak and style.prop_has_backup:
                try:
                    style.restore(cursor)
                except mEx.MultiError as e:
                    mLo.Lo.print("Write.style_left(): Unable to restore Property")
                    for err in e.errors:
                        mLo.Lo.print(f"  {err}")
                except Exception as e:
                    mLo.Lo.print("Write.style_left(): Unable to restore Property")
                    mLo.Lo.print(f"  {e}")

            _Events().trigger(WriteNamedEvent.STYLED, EventArgs.from_args(cargs))

        if pos <= 0:
            cursor.goRight(0, False)

    @overload
    @classmethod
    def style_left(cls, cursor: XTextCursor, pos: int, styles: Sequence[StyleT]) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Positions to style left.
            styles (Sequence[StyleT]): One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style_left(cls, cursor: XTextCursor, pos: int, prop_name: str, prop_val: object) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            pos (int): Positions to style left.
            prop_name (str): Property Name such as ``CharHeight``.
            prop_val (object): Property Value such as ``10``.

        Returns:
            None:
        """
        ...

    @classmethod
    def style_left(cls, *args, **kwargs) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor
            pos (int): Positions to style left
            styles (Sequence[StyleT]): One or more styles to apply to text.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

            Otherwise, the following events are triggered once.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-key-event`

        Returns:
            None:

        See Also:
            :py:meth:`~.Write.style`

        Note:
            This method restores the style properties to their original state after the style is applied.
            This is done so applied style properties are reset before next text is appended.
            This is not the case for :py:meth:`~.Write.style` method.

        .. versionchanged:: 0.9.0
            Added ``style_left(cursor: XTextCursor, pos: int, styles: Sequence[StyleT])`` overload.

            Added Events.
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cursor", "pos", "prop_name", "prop_val", "styles")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("style_left() got an unexpected keyword argument")
            ka[1] = kwargs.get("cursor", None)
            ka[2] = kwargs.get("pos", None)
            keys = ("prop_name", "styles")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("prop_val", None)
            return ka

        if count not in (3, 4):
            raise TypeError("style_left() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 4:
            return cls._style_left(kargs[1], kargs[2], kargs[3], kargs[4])
        return cls._style_left_style(kargs[1], kargs[2], kargs[3])

    # endregion style_left()

    @classmethod
    def dispatch_cmd_left(
        cls,
        vcursor: XTextViewCursor,
        pos: int,
        cmd: str,
        props: Iterable[PropertyValue] | None = None,
        frame: XFrame | None = None,
        toggle: bool = False,
    ) -> None:
        """
        Dispatches a command and applies it to selection based upon position.

        |lo_unsafe|

        Args:
            vcursor (XTextViewCursor): Text View Cursor.
            pos (int): Positions left to apply dispatch command.
            cmd (str): Dispatch command such as 'DefaultNumbering'.
            props (Iterable[PropertyValue], optional): properties for dispatch.
            frame (XFrame, optional): Frame to dispatch to.
            toggle (bool, optional): If True then dispatch will be preformed on selection.
                and again when deselected. Defaults to False.

        Returns:
            None:

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

        Attention:
            :py:meth:`Lo.dispatch_cmd <.utils.lo.Lo.dispatch_cmd>` method is called along with any of its events.
        """
        curr_pos = mSel.Selection.get_position(vcursor)
        vcursor.goLeft(curr_pos - pos, True)
        mLo.Lo.dispatch_cmd(cmd=cmd, props=props, frame=frame)  # type: ignore
        vcursor.goRight(curr_pos - pos, False)
        if toggle:
            mLo.Lo.dispatch_cmd(cmd=cmd, props=props, frame=frame)  # type: ignore

    # region    style_prev_paragraph()
    @staticmethod
    def _style_prev_paragraph_prop(cursor: XParagraphCursor, prop_val: object, prop_name: str) -> None:
        """Lo Safe Method."""
        cargs = KeyValCancelArgs("Write._style_prev_paragraph_prop", prop_name, prop_val)
        _Events().trigger(WriteNamedEvent.STYLE_PREV_PARA_PROP_SETTING, cargs)
        if cargs.cancel:
            return
        _Events().trigger(WriteNamedEvent.STYLING, cargs)
        if cargs.cancel:
            return
        old_val = mProps.Props.get(cursor, prop_name)

        cursor.gotoPreviousParagraph(True)  # select previous paragraph
        mProps.Props.set(cursor, **{prop_name: prop_val})

        # reset
        cursor.gotoNextParagraph(False)
        mProps.Props.set(cursor, **{prop_name: old_val})
        eargs = KeyValArgs.from_args(cargs)  # type: ignore
        _Events().trigger(WriteNamedEvent.STYLED, eargs)
        _Events().trigger(WriteNamedEvent.STYLE_PREV_PARA_PROP_SET, eargs)

    @classmethod
    def _style_prev_paragraph_style(cls, cursor: XTextCursor | XParagraphCursor, styles: Sequence[StyleT]) -> None:
        """Lo Safe Method."""
        if not styles:
            return
        c_styles_args = CancelEventArgs("Write._style_prev_paragraph_style")
        c_styles_args.event_data = styles
        _Events().trigger(WriteNamedEvent.STYLE_PREV_PARA_STYLES_SETTING, c_styles_args)
        if c_styles_args.cancel:
            return
        style_srv = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.drawing.FillProperties",
        )

        style_lst: List[StyleT] = []
        fill_lst: List[StyleT] = []
        style_data = cast(Sequence[StyleT], c_styles_args.event_data)
        for style in style_data:
            if not style.support_service(*style_srv):
                mLo.Lo.print(
                    f"style_prev_paragraph(), Supported services are {style_srv}. Not Supported style: {style}"
                )
                continue
            if FormatKind.TXT_CONTENT in style.prop_format_kind and FormatKind.PARA in style.prop_format_kind:
                fill_lst.append(style)
            else:
                style_lst.append(style)

        # has_prev = cursor.gotoPreviousParagraph(True)
        para_c = mLo.Lo.qi(XParagraphCursor, cursor)
        if para_c is None:
            para_c = cls.get_paragraph_cursor(cursor)

        if fill_lst:
            if has_prev := para_c.gotoPreviousParagraph(False):
                para_c.gotoEndOfParagraph(True)
                for style in fill_lst:
                    cargs = CancelEventArgs(c_styles_args.source)
                    cargs.event_data = style
                    _Events().trigger(WriteNamedEvent.STYLING, cargs)
                    if cargs.cancel:
                        continue
                    style.apply(para_c.TextParagraph)  # type: ignore
                    _Events().trigger(WriteNamedEvent.STYLED, EventArgs.from_args(cargs))

                para_c.gotoNextParagraph(False)

        if has_prev := para_c.gotoPreviousParagraph(False):
            if style_lst:
                para_c.gotoEndOfParagraph(True)
                for style in style_lst:
                    cargs = CancelEventArgs(c_styles_args.source)
                    cargs.event_data = style
                    _Events().trigger(WriteNamedEvent.STYLING, cargs)
                    if cargs.cancel:
                        continue
                    style.apply(para_c)
                    _Events().trigger(WriteNamedEvent.STYLED, EventArgs.from_args(cargs))

                para_c.gotoNextParagraph(False)

        e_style_args = EventArgs.from_args(c_styles_args)
        _Events().trigger(WriteNamedEvent.STYLE_PREV_PARA_STYLES_SET, e_style_args)

    @overload
    @classmethod
    def style_prev_paragraph(cls, cursor: XTextCursor, styles: Sequence[StyleT]) -> None:
        """
        Style previous paragraph.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            styles (Sequence[StyleT]): One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style_prev_paragraph(cls, cursor: XTextCursor, prop_val: Any) -> None:
        """
        Style previous paragraph.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            prop_val (Any): Property value.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def style_prev_paragraph(cls, cursor: XTextCursor, prop_val: Any, prop_name: str) -> None:
        """
        Style previous paragraph.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            prop_val (Any): Property value.
            prop_name (str): Property Name. Defaults to ``ParaStyleName``.

        Returns:
            None:
        """
        ...

    @classmethod
    def style_prev_paragraph(cls, *args, **kwargs) -> None:
        """
        Style previous paragraph.

        |lo_safe|

        Args:
            cursor (XTextCursor): Text Cursor.
            styles (Sequence[StyleT]): One or more styles to apply to text.
            prop_val (Any): Property value.
            prop_name (str): Property Name. Defaults to ``ParaStyleName``.

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

            Otherwise the following events are triggered once.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-key-event`

        Returns:
            None:

        .. collapse:: Example

            .. code-block:: python

                cursor = Write.get_cursor(doc)
                Write.style_prev_paragraph(cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust")

        .. versionchanged:: 0.9.0
            Added overload ``style_prev_paragraph(cursor: XTextCursor, styles: Sequence[StyleT])``
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cursor", "prop_name", "prop_val", "styles")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("style_prev_paragraph() got an unexpected keyword argument")
            ka[1] = kwargs.get("cursor", None)
            keys = ("prop_val", "styles")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("prop_name", None)
            return ka

        if count not in (2, 3):
            raise TypeError("style_prev_paragraph() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        # process code here

        if count == 3:
            cls._style_prev_paragraph_prop(cursor=kargs[1], prop_val=kargs[2], prop_name=kargs[3])
            return

        # count == 2
        arg2 = kargs[2]
        # arg2 must be string or Sequence[StyleT]
        # ParaStyleName can only be set a string value
        if isinstance(arg2, str):
            cls._style_prev_paragraph_prop(cursor=kargs[1], prop_val=arg2, prop_name="ParaStyleName")
        else:
            cls._style_prev_paragraph_style(cursor=kargs[1], styles=arg2)

    # endregion style_prev_paragraph()

    # endregion ---------- text cursor property methods ----------------

    # region ------------- style methods -------------------------------
    @staticmethod
    def create_style_para(text_doc: XTextDocument, style_name: str, styles: Sequence[StyleT] | None = None) -> XStyle:
        """
        Creates a paragraph style and adds it to document paragraph styles.

        |lo_unsafe|

        Args:
            text_doc (XTextDocument): Text Document.
            style_name (str): The name of the paragraph style.
            styles (Sequence[StyleT], optional): One or more styles to apply.

        Returns:
            XStyle: Newly created style.

        .. versionadded:: 0.9.2
        """
        para_styles = mInfo.Info.get_style_container(doc=text_doc, family_style_name="ParagraphStyles")

        # create new paragraph style properties set
        para_style = mLo.Lo.create_instance_msf(XStyle, "com.sun.star.style.ParagraphStyle", raise_err=True)
        props = mLo.Lo.qi(XPropertySet, para_style, raise_err=True)

        if styles:
            for style in styles:
                style.apply(props)

        # add the style to Document
        para_styles.insertByName(style_name, props)
        return mLo.Lo.qi(XStyle, para_style, True)

    @staticmethod
    def create_style_char(text_doc: XTextDocument, style_name: str, styles: Sequence[StyleT] | None = None) -> XStyle:
        """
        Creates a character style and adds it to document character styles.

        |lo_unsafe|

        Args:
            text_doc (XTextDocument): Text Document
            style_name (str): The name of the character style.
            styles (Sequence[StyleT], optional): One or more styles to apply.

        Returns:
            XStyle: Newly created style

        .. versionadded:: 0.9.2
        """
        char_styles = mInfo.Info.get_style_container(doc=text_doc, family_style_name="CharacterStyles")

        # create new character style properties set
        char_style = mLo.Lo.create_instance_msf(XStyle, "com.sun.star.style.CharacterStyle", raise_err=True)
        props = mLo.Lo.qi(XPropertySet, char_style, raise_err=True)

        if styles:
            for style in styles:
                style.apply(props)

        # add the style to Document
        char_styles.insertByName(style_name, props)
        return mLo.Lo.qi(XStyle, char_style, True)

    @staticmethod
    def get_page_text_width(text_doc: XTextDocument) -> int:
        """
        Get the width of the page's text area in ``1/100 mm`` units.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document

        Returns:
            int: Page Width in ``1/100 mm`` units on success; Otherwise 0
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            mLo.Lo.print("Could not access the standard page style")
            return 0

        try:
            width = int(props.getPropertyValue("Width"))
            left_margin = int(props.getPropertyValue("LeftMargin"))
            right_margin = int(props.getPropertyValue("RightMargin"))
            return width - (left_margin + right_margin)
        except Exception as e:
            mLo.Lo.print("Could not access standard page style dimensions")
            mLo.Lo.print(f"    {e}")
            return 0

    @staticmethod
    def get_page_text_size(text_doc: XTextDocument) -> Size:
        """
        Get page text size in ``1/100 mm`` units.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            ~ooodev.utils.data_type.size.Size: Page text Size in ``1/100 mm`` units.

        .. versionadded:: 0.9.0
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")
        try:
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            left_margin = int(props.getPropertyValue("LeftMargin"))
            right_margin = int(props.getPropertyValue("RightMargin"))
            top_margin = int(props.getPropertyValue("TopMargin"))
            btm_margin = int(props.getPropertyValue("BottomMargin"))
            text_width = width - (left_margin + right_margin)
            text_height = height - (top_margin + btm_margin)

            return Size(text_width, text_height)
        except Exception as e:
            raise Exception("Could not access standard page style dimensions") from e

    @staticmethod
    def get_page_size(text_doc: XTextDocument) -> Size:
        """
        Get page size in ``1/100 mm`` units.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            ~ooodev.utils.data_type.size.Size: Page Size in ``1/100 mm`` units.
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
    def set_page_format(text_doc: XTextDocument, paper_format: PaperFormat) -> bool:
        """
        Set Page Format.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document
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
            - :py:meth:`.set_a4_page_format`
        """
        cargs = CancelEventArgs(Write.set_page_format.__qualname__)
        cargs.event_data = {"paper_format": paper_format}
        _Events().trigger(WriteNamedEvent.PAGE_FORMAT_SETTING, cargs)
        if cargs.cancel:
            return False
        printable = mLo.Lo.qi(XPrintable, text_doc, True)
        printer_desc = mProps.Props.make_props(PaperFormat=paper_format)
        printable.setPrinter(printer_desc)
        _Events().trigger(WriteNamedEvent.PAGE_FORMAT_SET, EventArgs.from_args(cargs))
        return True

    @classmethod
    def set_a4_page_format(cls, text_doc: XTextDocument) -> bool:
        """
        Set Page Format to A4

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document

        Returns:
            bool: ``True`` if page format is set; Otherwise, ``False``

        See Also:
            :py:meth:`~.write.Write.set_page_format`

        Attention:
            :py:meth:`~.write.Write.set_page_format` method is called along with any of its events.
        """
        return cls.set_page_format(text_doc=text_doc, paper_format=PaperFormat.A4)

    # endregion ---------- style methods -------------------------------

    # region ------------- headers and footers -------------------------
    @classmethod
    def set_page_numbers(cls, text_doc: XTextDocument) -> XTextField:
        """
        Modify the footer via the page style for the document.
        Put page number & count in the center of the footer in Times New Roman, 12pt.

        |lo_unsafe|

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            PropertiesError: If unable to get properties
            Exception: If Unable to set page numbers

        Returns:
            XTextField: Page Number Field

        .. versionchanged:: 0.16.0
            Returns ``XTextField``.
        """
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")

        try:
            props.setPropertyValue("FooterIsOn", True)
            #   footer must be turned on in the document
            footer_text = mLo.Lo.qi(XText, props.getPropertyValue("FooterText"), True)
            footer_cursor = footer_text.createTextCursor()

            mProps.Props.set(
                footer_cursor,
                CharFontName=mInfo.Info.get_font_general_name(),
                CharHeight=12.0,
                ParaAdjust=ParagraphAdjust.CENTER,
            )

            # add text fields to the footer
            pg_number = cls.get_page_number()
            pg_content = mLo.Lo.qi(XTextContent, pg_number)
            if pg_content is None:
                raise mEx.MissingInterfaceError(
                    XTextContent, f"Missing interface for page number. {XTextContent.__pyunointerface__}"
                )
            cls._append_text_content(cursor=footer_cursor, text_content=pg_content)
            cls._append_text(cursor=footer_cursor, text=" of ")
            pg_count = cls.get_page_count()
            pg_count_content = mLo.Lo.qi(XTextContent, pg_count)
            if pg_count_content is None:
                raise mEx.MissingInterfaceError(
                    XTextContent, f"Missing interface for page count. {XTextContent.__pyunointerface__}"
                )
            cls._append_text_content(cursor=footer_cursor, text_content=pg_count_content)
            return pg_number
        except Exception as e:
            raise Exception("Unable to set page numbers") from e

    @staticmethod
    def get_page_number(
        numbering_type: NumberingTypeEnum = NumberingTypeEnum.ARABIC, sub_type: PageNumberType = PageNumberType.CURRENT
    ) -> XTextField:
        """
        Gets Arabic style number showing current page value.

        |lo_unsafe|

        Args:
            numbering_type (NumberingTypeEnum, optional): Numbering Type. Defaults to ``NumberingTypeEnum.ARABIC``.
            sub_type (PageNumberType, optional): Page Number Type. Defaults to ``PageNumberType.CURRENT``.

        Returns:
            XTextField: Page Number as Text Field

        See Also:
            :py:meth:`~.Write.get_current_page`
        """
        num_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageNumber", raise_err=True)
        mProps.Props.set(num_field, NumberingType=numbering_type.value, SubType=sub_type)
        return num_field

    @staticmethod
    def get_page_count(numbering_type: NumberingTypeEnum = NumberingTypeEnum.ARABIC) -> XTextField:
        """
        Return Arabic style number showing current page count/

        |lo_unsafe|

        Args:
            numbering_type (NumberingTypeEnum, optional): Numbering Type. Defaults to ``NumberingTypeEnum.ARABIC``.

        Returns:
            XTextField: Page Count as Text Field
        """
        pc_field = mLo.Lo.create_instance_msf(XTextField, "com.sun.star.text.TextField.PageCount", raise_err=True)
        mProps.Props.set(pc_field, NumberingType=numbering_type.value)
        return pc_field

    @staticmethod
    def _set_header_footer(
        text_doc: XTextDocument, text: str, kind: str = "h", styles: Sequence[StyleT] | None = None
    ) -> None:
        """Lo Safe Method."""
        props = mInfo.Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style container")
        try:
            # header or footer must be turned on in the document
            if kind == "h":
                props.setPropertyValue("HeaderIsOn", True)
                hf_text = mLo.Lo.qi(XText, props.getPropertyValue("HeaderText"), True)
            else:
                props.setPropertyValue("FooterIsOn", True)
                hf_text = mLo.Lo.qi(XText, props.getPropertyValue("FooterText"), True)
            hf_cursor = hf_text.createTextCursor()
            hf_cursor.gotoEnd(False)

            hf_props = mLo.Lo.qi(XPropertySet, hf_cursor, True)
            hf_props.setPropertyValue("CharFontName", mInfo.Info.get_font_general_name())
            hf_props.setPropertyValue("CharHeight", 10)
            hf_props.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT)
            txt_srv = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterPropertiesAsian",
                "com.sun.star.style.CharacterPropertiesComplex",
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphPropertiesAsian",
                "com.sun.star.style.ParagraphPropertiesComplex",
            )
            if styles is not None:
                for style in styles:
                    if style.support_service(*txt_srv):
                        style.apply(hf_props)

            hf_text.setString(f"{text}\n")
            f_kind = FormatKind.HEADER if kind == "h" else FormatKind.FOOTER
            if styles is not None:
                for style in styles:
                    if f_kind not in style.prop_format_kind:
                        continue
                    if FormatKind.DOC in style.prop_format_kind:
                        style.apply(text_doc)
                    else:
                        style.apply(props)
        except Exception as e:
            raise Exception("Unable to set header text") from e

    @classmethod
    def set_header(cls, text_doc: XTextDocument, text: str, styles: Sequence[StyleT] | None = None) -> None:
        """
        Modify the header via the page style for the document.
        Put the text on the right hand side in the header in
        a general font of 10pt.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.
            text (str): Header Text.
            styles (Sequence[StyleT]): Styles to apply to the text.

        Raises:
            PropertiesError: If unable to access properties.
            Exception: If unable to set header text.

        See Also:
            :py:meth:`~.write.Write.set_footer`

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_general_name`.

        .. versionchanged:: 0.9.2
            Added styles parameter
        """
        cls._set_header_footer(text_doc=text_doc, text=text, kind="h", styles=styles)

    @classmethod
    def set_footer(cls, text_doc: XTextDocument, text: str, styles: Sequence[StyleT] | None = None) -> None:
        """
        Modify the footer via the page style for the document.
        Put the text on the right hand side in the header in
        a general font of 10pt.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.
            text (str): Header Text.
            styles (Sequence[StyleT]): Styles to apply to the text.

        Raises:
            PropertiesError: If unable to access properties.
            Exception: If unable to set header text.

        See Also:
            :py:meth:`~.write.Write.set_header`

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_general_name`.

        .. versionadded:: 0.9.2
        """
        cls._set_header_footer(text_doc=text_doc, text=text, kind="f", styles=styles)

    @staticmethod
    def get_draw_page(text_doc: XTextDocument) -> XDrawPage:
        """
        Gets draw page.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            MissingInterfaceError: If text_doc does not implement XDrawPageSupplier interface.

        Returns:
            XDrawPage: Draw Page.
        """
        supp_page = mLo.Lo.qi(XDrawPageSupplier, text_doc, True)
        return supp_page.getDrawPage()

    @staticmethod
    def get_draw_pages(text_doc: XTextDocument) -> XDrawPages:
        """
        Gets draw pages.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            MissingInterfaceError: If text_doc does not implement XDrawPageSupplier interface.

        Returns:
            XDrawPages: Draw Page.
        """
        supp = mLo.Lo.qi(XDrawPagesSupplier, text_doc, True)
        return supp.getDrawPages()

    # endregion ---------- headers and footers -------------------------

    # region ------------- adding elements -----------------------------

    # region add_formula()
    @overload
    @classmethod
    def add_formula(cls, cursor: XTextCursor, formula: str) -> XTextContent:
        """
        Adds a formula.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Cursor.
            formula (str): formula.

        Returns:
            XTextContent: Embedded Object.
        """
        ...

    @overload
    @classmethod
    def add_formula(cls, cursor: XTextCursor, formula: str, styles: Sequence[StyleT]) -> XTextContent:
        """
        Adds a formula.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Cursor.
            formula (str): formula.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Returns:
            XTextContent: Embedded Object.
        """
        ...

    @classmethod
    def add_formula(cls, cursor: XTextCursor, formula: str, styles: Sequence[StyleT] | None = None) -> XTextContent:
        """
        Adds a formula.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Cursor.
            formula (str): formula.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If unable to create text.TextEmbeddedObject.
            CancelEventError: If event ``WriteNamedEvent.FORMULA_ADDING`` is cancelled.
            Exception: If unable to add formula.

        Returns:
            XTextContent: Embedded Object.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.obj </src/format/ooodev.format.writer.direct.obj>`

        Note:
            Event args ``event_data`` is a dictionary containing ``formula`` and ``cursor``.

        .. versionchanged:: 0.9.0
            Now returns the embedded Object instead of bool value.
            Added style parameter that allows for all styles that support ``com.sun.star.text.TextEmbeddedObject`` service.
        """
        # sourcery skip: raise-specific-error
        result = None
        cargs = CancelEventArgs(Write.add_formula.__qualname__)
        cargs.event_data = {"cursor": cursor, "formula": formula}
        _Events().trigger(WriteNamedEvent.FORMULA_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        formula = cast(str, cargs.event_data["formula"])
        embed_content = mLo.Lo.create_instance_msf(
            XTextContent, "com.sun.star.text.TextEmbeddedObject", raise_err=True
        )
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
            result = embed_content
            if styles:
                # srv = ("com.sun.star.text.TextEmbeddedObject",)
                for style in styles:
                    # if style.support_service(*srv):
                    style.apply(embed_content)
            mLo.Lo.print(f'Inserted formula "{formula}"')
        except Exception as e:
            raise Exception(f'Insertion fo formula "{formula}" failed:') from e
        _Events().trigger(WriteNamedEvent.FORMULA_ADDED, EventArgs.from_args(cargs))
        return result

    # endregion add_formula()

    @classmethod
    def add_hyperlink(cls, cursor: XTextCursor, label: str, url_str: str) -> bool:
        """
        Add a hyperlink.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            label (str): Hyperlink label.
            url_str (str): Hyperlink url.

        Raises:
            CreateInstanceMsfError: If unable to create TextField.URL instance.
            Exception: If unable to create hyperlink.

        Returns:
            bool: ``True`` if hyperlink is added; Otherwise, ``False``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``label``, ``url_str`` and ``cursor``.
        """
        cargs = CancelEventArgs(Write.add_hyperlink.__qualname__)
        cargs.event_data = {"cursor": cursor, "label": label, "url_str": url_str}
        _Events().trigger(WriteNamedEvent.HYPER_LINK_ADDING, cargs)
        if cargs.cancel:
            return False

        label = cast(str, cargs.event_data["label"])
        url_str = cast(str, cargs.event_data["url_str"])

        try:
            link = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextField.URL")
            if link is None:
                raise ValueError("Null Value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.text.TextField.URL") from e
        try:
            mProps.Props.set(link, URL=url_str, Representation=label)

            cls._append_text_content(cursor, link)
            mLo.Lo.print("Added hyperlink")
        except Exception as e:
            raise Exception("Unable to add hyperlink") from e
        _Events().trigger(WriteNamedEvent.HYPER_LINK_ADDED, EventArgs.from_args(cargs))
        return True

    @classmethod
    def add_bookmark(cls, cursor: XTextCursor, name: str) -> bool:
        """
        Adds bookmark.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            name (str): Bookmark name.

        Returns:
            bool: ``True`` if bookmark is added; Otherwise, ``False``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``name`` and ``cursor``.
        """
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(Write.add_bookmark.__qualname__)
        cargs.event_data = {"cursor": cursor, "name": name}
        _Events().trigger(WriteNamedEvent.BOOKMARK_ADDING, cargs)
        if cargs.cancel:
            return False

        name = cast(str, cargs.event_data["name"])

        bmk_content = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.Bookmark")
        if bmk_content is None:
            mLo.Lo.print(f'Unable to create Bookmark instance for "{name}"')
            return False
        try:
            bmk_named = mLo.Lo.qi(XNamed, bmk_content, True)
            bmk_named.setName(name)

            cls._append_text_content(cursor, bmk_content)
        except Exception as e:
            mLo.Lo.print(f"Unable to add bookmark '{name}'")
            mLo.Lo.print(f"  {e}")
            return False
        _Events().trigger(WriteNamedEvent.BOOKMARK_ADDED, EventArgs.from_args(cargs))
        return True

    @staticmethod
    def find_bookmark(text_doc: XTextDocument, bm_name: str) -> XTextContent | None:
        """
        Finds a bookmark.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.
            bm_name (str): Bookmark name.

        Raises:
            MissingInterfaceError: if text_doc does not implement XBookmarksSupplier interface.

        Returns:
            XTextContent | None: Bookmark if found; Otherwise, ``None``.
        """
        supplier = mLo.Lo.qi(XBookmarksSupplier, text_doc, True)

        named_bookmarks = supplier.getBookmarks()
        bookmark = None

        try:
            bookmark = named_bookmarks.getByName(bm_name)
        except Exception:
            mLo.Lo.print(f"Bookmark '{bm_name}' not found")
            return None
        return mLo.Lo.qi(XTextContent, bookmark)

    @classmethod
    def add_text_frame(
        cls,
        *,
        cursor: XTextCursor,
        text: str = "",
        ypos: int | UnitT = 300,
        width: int | UnitT = 5000,
        height: int | UnitT = 5000,
        page_num: int = 1,
        border_color: Color | None = None,
        background_color: Color | None = None,
        styles: Sequence[StyleT] | None = None,
    ) -> XTextFrame:
        """
        Adds a text frame.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            text (str, optional): Frame Text.
            ypos (int, UnitT. optional): Frame Y pos in ``1/100th mm`` or :ref:`proto_unit_obj`. Default ``300``.
            width (int, UnitT, optional): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT, optional): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.
            page_num (int, optional): Page Number to add text frame. If ``0`` Then Frame is anchored to paragraph. Default ``1``.
            border_color (:py:data:`~.utils.color.Color`, optional):.color.Color`, optional): Border Color.
            background_color (:py:data:`~.utils.color.Color`, optional): Background Color.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextFrame`` service are applied.

        Raises:
            CreateInstanceMsfError: If unable to create text.TextFrame.
            CancelEventError: If ``WriteNamedEvent.TEXT_FRAME_ADDING`` event is cancelled.
            Exception: If unable to add text frame.

        Returns:
            XTextFrame: Text frame that is added to document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TEXT_FRAME_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TEXT_FRAME_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.writer.direct.frame </src/format/ooodev.format.writer.direct.frame>` subpackages.

        Note:
            Event args ``event_data`` is a dictionary containing all method args.

        See Also:
            - :py:class:`~.utils.color.CommonColor`
            - :py:class:`~.utils.color.StandardColor`

        .. versionchanged:: 0.9.0
            Now returns the added text frame instead of bool value.
            Added ``UnitMM`` values.
            ``border_color`` and ``background_color`` now default to ``None``.
            Added style parameter that allows for all styles that support ``com.sun.star.text.TextFrame`` service.
        """
        # sourcery skip: raise-specific-error

        # If Text is added to both frames that are to be chained together then
        # LO will not chain them.
        # The Frame.ChainNextName and Frame.ChainPrevName properties cannot be set. and do not raise any error.
        # This is the default behavior and makes sense.
        # If the frame to flow to has text already then previous frame cannot flow to it.
        #
        # xframe = Lo.create_instance_msf(XTextFrame, "com.sun.star.text.ChainedTextFrame")
        # Raises error: see, https://bugs.documentfoundation.org/show_bug.cgi?id=153825
        # tf_shape = Lo.qi(XShape, xframe, True)

        cargs = CancelEventArgs(Write.add_text_frame.__qualname__)
        cargs.event_data = {
            "cursor": cursor,
            "ypos": ypos,
            "text": text,
            "width": width,
            "height": height,
            "page_num": page_num,
            "border_color": border_color,
            "background_color": background_color,
        }
        _Events().trigger(WriteNamedEvent.TEXT_FRAME_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        arg_ypos = cast(Union[int, UnitT], cargs.event_data["ypos"])
        text = cast(str, cargs.event_data["text"])
        arg_width = cast(Union[int, UnitT], cargs.event_data["width"])
        arg_height = cast(Union[int, UnitT], cargs.event_data["height"])
        page_num = cast(int, cargs.event_data["page_num"])
        border_color = cast(Union[Color, None], cargs.event_data["border_color"])
        background_color = cast(Union[Color, None], cargs.event_data["background_color"])

        try:
            ypos = arg_ypos.get_value_mm100()  # type: ignore
        except AttributeError:
            ypos = int(arg_ypos)  # type: ignore
        try:
            width = arg_width.get_value_mm100()  # type: ignore
        except AttributeError:
            width = int(arg_width)  # type: ignore
        try:
            height = arg_height.get_value_mm100()  # type: ignore
        except AttributeError:
            height = int(arg_height)  # type: ignore

        # if mLo.Lo.bridge_connector.headless:
        #     # this does not allow chaining. See above comments.
        #     xframe = mLo.Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame", raise_err=True)
        # else:
        #     xframe = cls._add_text_frame_via_dispatch(ypos=ypos, width=width, height=height)

        xframe = mLo.Lo.create_instance_msf(XTextFrame, "com.sun.star.text.TextFrame", raise_err=True)

        try:
            tf_shape = mLo.Lo.qi(XShape, xframe, True)

            # set dimensions of the text frame
            tf_shape.setSize(UnoSize(width, height))

            #  anchor the text frame
            frame_props = mLo.Lo.qi(XPropertySet, xframe, True)
            # if page number is Not include for TextContentAnchorType.AT_PAGE
            # then Lo Default so At AT_PARAGRAPH
            if not page_num or page_num < 1:
                frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH)
            else:
                frame_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PAGE)
                frame_props.setPropertyValue("AnchorPageNo", page_num)

            frame_props.setPropertyValue("FrameIsAutomaticHeight", True)  # will grow if necessary

            # add a red border around all 4 sides
            border = BorderLine()
            border.OuterLineWidth = 1
            if border_color is not None:
                border.Color = border_color  # type: ignore

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

            # insert text frame into document (order is important here)
            cls._append_text_content(cursor, xframe)
            cls.end_paragraph(cursor)

            if text:
                xframe_text = xframe.getText()
                xtext_range = mLo.Lo.qi(XTextRange, xframe_text.createTextCursor(), True)
                xframe_text.insertString(xtext_range, text, False)

            if styles:
                # srv = ("com.sun.star.text.TextFrame", "com.sun.star.text.ChainedTextFrame")
                for style in styles:
                    # if style.support_service(*srv):
                    style.apply(xframe)
            # add text into the text frame
        except Exception as e:
            raise Exception("Insertion of text frame failed:") from e
        _Events().trigger(WriteNamedEvent.TEXT_FRAME_ADDED, EventArgs.from_args(cargs))
        return xframe

    @classmethod
    def _add_text_frame_via_dispatch(cls, ypos: int, width: int, height: int) -> XTextFrame:
        """Lo UN-Safe Method."""

        # this method is not currently being used.
        # It works so it is left here for possible future use.
        def filter_frame(val: str) -> bool:
            regex = r"Frame\d+$"
            return bool(re.match(regex, val))

        args = {"AnchorType": 0, "Pos.X": 1000, "Pos.Y": ypos, "Size.Width": width, "Size.Height": height}
        vals = mProps.Props.make_props(**args)
        mLo.Lo.dispatch_cmd(cmd="Escape")
        mLo.Lo.delay(200)
        mLo.Lo.dispatch_cmd(cmd="InsertFrame", props=vals)
        mLo.Lo.delay(200)
        mLo.Lo.dispatch_cmd(cmd="Escape")
        mLo.Lo.delay(200)
        frames = cls.get_text_frames(cls.active_doc)
        if frames is None:
            raise RuntimeError("Failed to get frames from document")
        names = frames.getElementNames()
        if not names:
            raise RuntimeError("Failed to add frames to document")

        if len(names) == 1:
            return frames.getByName(names[0])

        # Filter out any frames that don't start with Frame...
        filtered_names = filter(filter_frame, names)
        frame_names = list(filtered_names)
        # Sort in human readable so the highest frame name is at the end of the list.
        frame_names.sort(key=mUtil.Util.natural_key_sorter)
        return frames.getByName(frame_names.pop())

    @classmethod
    def add_table(
        cls,
        cursor: XTextCursor,
        table_data: Table,
        header_bg_color: Color | None = CommonColor.DARK_BLUE,
        header_fg_color: Color | None = CommonColor.WHITE,
        tbl_bg_color: Color | None = CommonColor.LIGHT_BLUE,
        tbl_fg_color: Color | None = CommonColor.BLACK,
        first_row_header: bool = True,
        styles: Sequence[StyleT] | None = None,
    ) -> XTextTable:
        """
        Adds a table.

        Each row becomes a row of the table. The first row is treated as a header.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            table_data (Table): 2D Table with the the first row containing column names.
            header_bg_color (:py:data:`~.utils.color.Color`, optional): Table header background color.
                Set to None to ignore header color. Defaults to ``CommonColor.DARK_BLUE``.
            header_fg_color (:py:data:`~.utils.color.Color`, optional): Table header foreground color.
                Set to None to ignore header color. Defaults to ``CommonColor.WHITE``.
            tbl_bg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.LIGHT_BLUE``.
            tbl_fg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.BLACK``.
            first_row_header (bool, optional): If ``True`` First row is treated as header data. Default ``True``.
            styles (Sequence[StyleT], optional): One or more styles to apply to frame.
                Only styles that support ``com.sun.star.text.TextTable`` service are applied.

        Raises:
            ValueError: If table_data is empty.
            CreateInstanceMsfError: If unable to create instance of text.TextTable.
            CancelEventError:  If ``WriteNamedEvent.TABLE_ADDING`` event cancelled.
            Exception: If unable to add table.

        Returns:
            XTextTable: Table that is added to document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method args.

        See Also:
            - :ref:`help_writer_format_direct_table`
            - :py:class:`~.utils.color.CommonColor`
            - :py:meth:`~.utils.table_helper.TableHelper.table_2d_to_dict`
            - :py:meth:`~.utils.table_helper.TableHelper.table_dict_to_table`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.writer.direct.table </src/format/ooodev.format.writer.direct.table>` subpackages.

        .. versionchanged:: 0.9.0
            Now returns added table instead of bool value.
            Added options ``first_row_header`` and ``styles``.
        """
        # sourcery skip: remove-unnecessary-cast

        cargs = CancelEventArgs(Write.add_table.__qualname__)
        cargs.event_data = {
            "cursor": cursor,
            "table_data": table_data,
            "header_bg_color": header_bg_color,
            "header_fg_color": header_fg_color,
            "tbl_bg_color": tbl_bg_color,
            "tbl_fg_color": tbl_fg_color,
            "first_row_header": first_row_header,
            "styles": styles,
        }
        _Events().trigger(WriteNamedEvent.TABLE_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        header_bg_color = cast(Union[Color, None], cargs.event_data["header_bg_color"])
        header_fg_color = cast(Union[Color, None], cargs.event_data["header_fg_color"])
        tbl_bg_color = cast(Union[Color, None], cargs.event_data["tbl_bg_color"])
        tbl_fg_color = cast(Union[Color, None], cargs.event_data["tbl_fg_color"])
        first_row_header = cast(bool, cargs.event_data["first_row_header"])

        def make_cell_name(row: int, col: int) -> str:
            return TableHelper.make_cell_name(row=row + 1, col=col + 1)

        def set_cell_header(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if first_row_header and header_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                mProps.Props.set(text_cursor, CharColor=header_fg_color)

            cell_text.setString(str(data))

        def set_cell_text(cell_name: str, data: str, table: XTextTable) -> None:
            cell_text = mLo.Lo.qi(XText, table.getCellByName(cell_name), True)
            if first_row_header is False or tbl_fg_color is not None:
                text_cursor = cell_text.createTextCursor()
                props = {}
                if not first_row_header:
                    # By default the first row has a style by the name of: Table Heading
                    # Table Contents is the default for cell that are not in the header row.
                    props["ParaStyleName"] = "Table Contents"
                if tbl_fg_color is not None:
                    props["CharColor"] = tbl_fg_color
                mProps.Props.set(text_cursor, **props)

            cell_text.setString(str(data))

        num_rows = len(table_data)
        if num_rows == 0:
            raise ValueError("table_data has no values")
        try:
            table = mLo.Lo.create_instance_msf(XTextTable, "com.sun.star.text.TextTable")
            if table is None:
                raise ValueError("Null Value")
        except Exception as e:
            raise mEx.CreateInstanceMsfError(XTextTable, "com.sun.star.text.TextTable") from e

        try:
            num_cols = len(table_data[0])
            if num_cols == 0:
                raise ValueError("table_data has no columns")
            mLo.Lo.print(f"Creating table rows: {num_rows}, cols: {num_cols}")
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
            if first_row_header and header_bg_color is not None:
                rows = table.getRows()
                mProps.Props.set(rows.getByIndex(0), BackColor=header_bg_color)

            #  write table header
            if first_row_header:
                row_data = table_data[0]
                for x in range(num_cols):
                    set_cell_header(make_cell_name(0, x), row_data[x], table)
                    # e.g. "A1", "B1", "C1", etc

                # insert table body
                for y in range(1, num_rows):  # start in 2nd row
                    row_data = table_data[y]
                    for x in range(num_cols):
                        set_cell_text(make_cell_name(y, x), row_data[x], table)
            else:
                # insert table body
                for y in range(num_rows):  # start in 1st row
                    row_data = table_data[y]
                    for x in range(num_cols):
                        set_cell_text(make_cell_name(y, x), row_data[x], table)

            if styles:
                srv = ("com.sun.star.text.TextTable",)
                for style in styles:
                    if style.support_service(*srv):
                        style.apply(table)
        except Exception as e:
            raise Exception("Table insertion failed:") from e
        _Events().trigger(WriteNamedEvent.TABLE_ADDED, EventArgs.from_args(cargs))
        return table

    # region    add_image_link()

    @overload
    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr) -> XTextContent:
        """
        Add Image Link.

        |lo_unsafe|

        Args:
            doc (XTextDocument): Text Document.
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    # style: Sequence[StyleT] = None
    @overload
    @classmethod
    def add_image_link(
        cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, *, width: int | UnitT, height: int | UnitT
    ) -> XTextContent:
        """
        Add Image Link.

        |lo_unsafe|

        Args:
            doc (XTextDocument): Text Document.
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def add_image_link(
        cls,
        doc: XTextDocument,
        cursor: XTextCursor,
        fnm: PathOrStr,
        *,
        styles: Sequence[StyleT],
    ) -> XTextContent:
        """
        Add Image Link.

        |lo_unsafe|

        Args:
            doc (XTextDocument): Text Document.
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def add_image_link(
        cls,
        doc: XTextDocument,
        cursor: XTextCursor,
        fnm: PathOrStr,
        *,
        width: int | UnitT,
        height: int | UnitT,
        styles: Sequence[StyleT],
    ) -> XTextContent:
        """
        Add Image Link.

        |lo_unsafe|

        Args:
            doc (XTextDocument): Text Document.
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @classmethod
    def add_image_link(
        cls,
        doc: XTextDocument,
        cursor: XTextCursor,
        fnm: PathOrStr,
        *,
        width: int | UnitT = 0,
        height: int | UnitT = 0,
        styles: Sequence[StyleT] | None = None,
    ) -> XTextContent:
        """
        Add Image Link.

        |lo_unsafe|

        Args:
            doc (XTextDocument): Text Document.
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If Unable to create text.TextGraphicObject.
            MissingInterfaceError: If unable to obtain XPropertySet interface.
            Exception: If unable to add image.
            CancelEventError: If ``IMAGE_LINK_ADDING`` event is canceled.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.image </src/format/ooodev.format.writer.direct.image>`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.

        .. versionchanged:: 0.16.0
            Raises CancelEventError if event is canceled.

        .. versionchanged:: 0..0
            Return image shape instead of boolean.
        """
        # see Also: https://ask.libreoffice.org/t/graphicurl-no-longer-works-in-6-1-0-3/35459/3
        # see Also: https://tomazvajngerl.blogspot.com/2018/03/improving-image-handling-in-libreoffice.html
        result = None
        cargs = CancelEventArgs(Write.add_image_link.__qualname__)
        cargs.event_data = {
            "doc": doc,
            "cursor": cursor,
            "fnm": fnm,
            "width": width,
            "height": height,
        }
        _Events().trigger(WriteNamedEvent.IMAGE_LINK_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        fnm = cast("PathOrStr", cargs.event_data["fnm"])
        width = cast(Union[int, UnitT], cargs.event_data["width"])
        height = cast(Union[int, UnitT], cargs.event_data["height"])

        tgo = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextGraphicObject", raise_err=True)
        props = mLo.Lo.qi(XPropertySet, tgo, True)

        try:
            graphic = mImgLo.ImagesLo.load_graphic_file(fnm)
            props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER)
            props.setPropertyValue("Graphic", graphic)

            # optionally set the width and height
            if width is not None:
                try:
                    props.setPropertyValue("Width", width.get_value_mm100())  # type: ignore
                except AttributeError:
                    props.setPropertyValue("Width", int(width))  # type: ignore
            if height is not None:
                try:
                    props.setPropertyValue("Height", height.get_value_mm100())  # type: ignore
                except AttributeError:
                    props.setPropertyValue("Height", int(height))  # type: ignore

            # append image to document
            cls._append_text_content(cursor, tgo)
            # end the paragraph.
            cls.end_line(cursor)
            # set any styles for the image.
            if styles:
                # is is important for some format styles such as Crop that styles
                # not be applied until after they have been added to the document.
                # srv = ("com.sun.star.text.TextGraphicObject",)
                for style in styles:
                    # if style.support_service(*srv):
                    style.apply(tgo)
            result = tgo
        except Exception as e:
            raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e
        _Events().trigger(WriteNamedEvent.IMAGE_LINK_ADDED, EventArgs.from_args(cargs))
        return result

    # endregion add_image_link()

    # region    add_image_shape()
    @overload
    @classmethod
    def add_image_shape(cls, cursor: XTextCursor, fnm: PathOrStr) -> XShape:
        """
        Add Image Shape.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.

        Returns:
            XShape: Image Shape on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def add_image_shape(cls, cursor: XTextCursor, fnm: PathOrStr, width: int | UnitT, height: int | UnitT) -> XShape:
        """
        Add Image Shape.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            XShape: Image Shape on success; Otherwise, ``None``.
        """
        ...

    @classmethod
    def add_image_shape(
        cls, cursor: XTextCursor, fnm: PathOrStr, width: int | UnitT = 0, height: int | UnitT = 0
    ) -> XShape:
        """
        Add Image Shape.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.

        Raises:
            CreateInstanceMsfError: If unable to create drawing.GraphicObjectShape.
            ValueError: If unable to get image.
            MissingInterfaceError: If require interface cannot be obtained.
            Exception: If unable to add image shape.
            CancelEventError: if ``IMAGE_SHAPE_ADDING`` event is canceled.

        Returns:
            XShape: Image Shape on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.

        .. versionchanged:: 0.16.0
            Raises CancelEventError if event is canceled.

        .. versionchanged:: 0.9.0
            Return image shape instead of boolean.
        """
        result = None
        cargs = CancelEventArgs(Write.add_image_shape.__qualname__)
        cargs.event_data = {
            "cursor": cursor,
            "fnm": fnm,
            "width": width,
            "height": height,
        }
        _Events().trigger(WriteNamedEvent.IMAGE_SHAPE_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        fnm = cast("PathOrStr", cargs.event_data["fnm"])
        width = cast(Union[int, UnitT], cargs.event_data["width"])
        height = cast(Union[int, UnitT], cargs.event_data["height"])

        pth = mFileIO.FileIO.get_absolute_path(fnm)

        try:
            size_set = False
            im_size = None
            if width is not None and height is not None:
                try:
                    w = width.get_value_mm100()  # type: ignore
                except AttributeError:
                    w = int(width)  # type: ignore

                try:
                    h = height.get_value_mm100()  # type: ignore
                except AttributeError:
                    h = int(height)  # type: ignore

                if w > 0 and h > 0:
                    im_size = Size(w, h)
                    size_set = True
            if not size_set:
                im_size = mImgLo.ImagesLo.get_size_100mm(pth)  # in 1/100 mm units
                if im_size is None:
                    raise ValueError(f"Unable to get image from {pth}")

            # create TextContent for an empty graphic
            gos = mLo.Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.GraphicObjectShape")
            if gos is None:
                raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.drawing.GraphicObjectShape")

            bitmap = mImgLo.ImagesLo.get_bitmap(pth)
            if bitmap is None:
                raise ValueError(f"Unable to get bitmap of {pth}")
            # store the image's bitmap as the graphic shape's URL's value
            mProps.Props.set(gos, GraphicURL=bitmap)

            # set the shape's size
            xdraw_shape = mLo.Lo.qi(XShape, gos, True)
            if im_size is None:
                raise ValueError(f"Unable to get image size of {pth}")
            xdraw_shape.setSize(im_size.get_uno_size())

            # insert image shape into the document, followed by newline
            cls._append_text_content(cursor, gos)
            cls.end_line(cursor)
            result = xdraw_shape
        except ValueError:
            raise
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e
        _Events().trigger(WriteNamedEvent.IMAGE_SHAPE_ADDED, EventArgs.from_args(cargs))
        return result

    # endregion add_image_shape()

    @classmethod
    def add_line_divider(cls, cursor: XTextCursor, line_width: int) -> None:
        """
        Adds a line divider.

        |lo_unsafe|

        Args:
            cursor (XTextCursor): Text Cursor.
            line_width (int): Line width.

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
            line_shape.setSize(UnoSize(line_width, 0))

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

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            Exception: If unable to get text graphics.

        Returns:
            List[XGraphic]: Text Graphics.

        Note:
            If there is error getting a graphic link then it is ignored
            and not added to the return value.
        """
        try:
            name_access = cls.get_graphic_links(text_doc)
            if name_access is None:
                raise ValueError("Unable to get Graphic Links")
            names = name_access.getElementNames()

            pics: List[XGraphic] = []
            for name in names:
                graphic_link = None
                with contextlib.suppress(UnoException):
                    graphic_link = name_access.getByName(name)
                if graphic_link is None:
                    mLo.Lo.print(f"No graphic found for {name}")
                else:
                    try:
                        graphic = mImgLo.ImagesLo.load_graphic_link(graphic_link)
                        if graphic is not None:
                            pics.append(graphic)
                    except Exception as e:
                        mLo.Lo.print(f"{name} could not be accessed:")
                        mLo.Lo.print(f"    {e}")
            return pics
        except Exception as e:
            raise Exception("Get text graphics failed:") from e

    @staticmethod
    def get_graphic_links(doc: XComponent) -> XNameAccess | None:
        """
        Gets graphic links.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            MissingInterfaceError: if doc does not implement ``XTextGraphicObjectsSupplier`` interface.

        Returns:
            XNameAccess | None: Graphic Links on success, Otherwise, ``None``.
        """
        ims_supplier = mLo.Lo.qi(XTextGraphicObjectsSupplier, doc, True)

        name_access = ims_supplier.getGraphicObjects()
        if name_access is None:
            mLo.Lo.print("Name access to graphics not possible")
            return None

        if not name_access.hasElements():
            mLo.Lo.print("No graphics elements found")
            return None

        return name_access

    @staticmethod
    def get_text_frames(doc: XComponent) -> XNameAccess | None:
        """
        Gets document Text Frames.

        |lo_safe|

        Args:
            doc (XComponent): Document

        Raises:
            MissingInterfaceError: if doc does not implement ``XTextFramesSupplier`` interface

        Returns:
            XNameAccess | None: Text Frames on success, Otherwise, None

        .. versionadded:: 0.9.0
        """
        supplier = mLo.Lo.qi(XTextFramesSupplier, doc, True)

        name_access = supplier.getTextFrames()
        if name_access is None:
            mLo.Lo.print("Name access to text frames not possible")
            return None

        if not name_access.hasElements():
            mLo.Lo.print("No text frame elements found")
            return None

        return name_access

    @staticmethod
    def is_anchored_graphic(graphic: Any) -> bool:
        """
        Gets if a graphic object is an anchored graphic.

        |lo_safe|

        Args:
            graphic (Any): object that implements XServiceInfo.

        Returns:
            bool: ``True`` if is anchored graphic; Otherwise, ``False``.
        """
        service_info = mLo.Lo.qi(XServiceInfo, graphic)
        return (
            service_info is not None
            and service_info.supportsService("com.sun.star.text.TextContent")
            and service_info.supportsService("com.sun.star.text.TextGraphicObject")
        )

    @staticmethod
    def get_shapes(text_doc: XTextDocument) -> XDrawPage:
        """
        Gets shapes.

        |lo_safe|

        Args:
            text_doc (XTextDocument): Text Document.

        Raises:
            MissingInterfaceError: If text_doc does not implement ``XDrawPageSupplier`` interface.

        Returns:
            XDrawPage: shapes
        """
        draw_page_supplier = mLo.Lo.qi(XDrawPageSupplier, text_doc, True)

        return draw_page_supplier.getDrawPage()

    # endregion ---------- extracting graphics from text doc -----------

    # region ------------  Linguistic API ------------------------------

    @overload
    @classmethod
    def print_services_info(cls, lingo_mgr: XLinguServiceManager2) -> None:
        """
        Prints service info to console

        Args:
            lingo_mgr (XLinguServiceManager2): Service manager
        """
        ...

    @overload
    @classmethod
    def print_services_info(cls, lingo_mgr: XLinguServiceManager2, loc: Locale) -> None:
        """
        Prints service info to console

        Args:
            lingo_mgr (XLinguServiceManager2): Service manager
            loc (Locale) : Locale
        """
        ...

    @classmethod
    def print_services_info(cls, lingo_mgr: XLinguServiceManager2, loc: Locale | None = None) -> None:
        """
        Prints service info to console.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service manager.
            loc (Locale | None, Optional) : Locale. Default ``Locale("en", "US", "")``.
        """
        if loc is None:
            loc = Locale("en", "US", "")
        print("Available Services:")
        cls.print_avail_service_info(lingo_mgr, "SpellChecker", loc)
        cls.print_avail_service_info(lingo_mgr, "Thesaurus", loc)
        cls.print_avail_service_info(lingo_mgr, "Hyphenator", loc)
        cls.print_avail_service_info(lingo_mgr, "Proofreader", loc)
        print()

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
        Prints available service info to console.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manger.
            service (str): Service Name.
            loc (Locale): Locale.
        """
        service_names = lingo_mgr.getAvailableServices(f"com.sun.star.linguistic2.{service}", loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")

    @staticmethod
    def print_config_service_info(lingo_mgr: XLinguServiceManager2, service: str, loc: Locale) -> None:
        """
        Print config service info to console.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager.
            service (str): Service Name.
            loc (Locale): Locale.
        """
        service_names = lingo_mgr.getConfiguredServices(f"com.sun.star.linguistic2.{service}", loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")

    @staticmethod
    def print_locales(service: str, loc: Iterable[Locale]) -> None:
        """
        Print locales to console.

        |lo_safe|

        Args:
            service (str): Service.
            loc (Iterable[Locale]): Locale's.
        """
        countries: List[str] = []
        for l in loc:
            if c_str := l.Country.strip():
                countries.append(c_str)
        countries.sort()

        print(f"Locales for {service} ({len(countries)})")
        for i, country in enumerate(countries):
            # print 10 per line
            if (i > 9) and (i % 10) == 0:
                print()
            print(f"  {country}", end="")
        print()
        print()

    @overload
    @staticmethod
    def set_configured_services(lingo_mgr: XLinguServiceManager2, service: str, impl_name: str) -> bool:
        """
        Set configured Services.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager.
            service (str): Service Name.
            impl_name (str): Service implementation name.

        Returns:
            bool: ``True`` if ``CONFIGURED_SERVICES_SETTING`` event is not canceled; Otherwise, ``False``.
        """
        ...

    @overload
    @staticmethod
    def set_configured_services(lingo_mgr: XLinguServiceManager2, service: str, impl_name: str, loc: Locale) -> bool:
        """
        Set configured Services.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager.
            service (str): Service Name.
            impl_name (str): Service implementation name.
            loc (Locale): Local used to spell words.

        Returns:
            bool: ``True`` if ``CONFIGURED_SERVICES_SETTING`` event is not canceled; Otherwise, ``False``.
        """
        ...

    @staticmethod
    def set_configured_services(
        lingo_mgr: XLinguServiceManager2, service: str, impl_name: str, loc: Locale | None = None
    ) -> bool:
        """
        Set configured Services.

        |lo_safe|

        Args:
            lingo_mgr (XLinguServiceManager2): Service Manager
            service (str): Service Name
            impl_name (str): Service implementation name
            loc (Locale | None, optional): Local used to spell words. Default ``Locale("en", "US", "")``

        Returns:
            bool: ``True`` if ``CONFIGURED_SERVICES_SETTING`` event is not canceled; Otherwise, ``False``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.CONFIGURED_SERVICES_SETTING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.CONFIGURED_SERVICES_SET` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.
        """
        cargs = CancelEventArgs(Write.set_configured_services.__qualname__)
        cargs.event_data = {
            "lingo_mgr": lingo_mgr,
            "service": service,
            "impl_name": impl_name,
        }
        _Events().trigger(WriteNamedEvent.CONFIGURED_SERVICES_SETTING, cargs)
        if cargs.cancel:
            return False
        service = cast(str, cargs.event_data["service"])
        impl_name = cast(str, cargs.event_data["impl_name"])
        if loc is None:
            loc = Locale("en", "US", "")
        impl_names = (impl_name,)
        lingo_mgr.setConfiguredServices(f"com.sun.star.linguistic2.{service}", loc, impl_names)
        _Events().trigger(WriteNamedEvent.CONFIGURED_SERVICES_SET, EventArgs.from_args(cargs))
        return True

    @classmethod
    def dicts_info(cls) -> None:
        """
        Prints dictionary info to console.

        |lo_unsafe|
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
        Prints dictionaries info to console.

        |lo_safe|

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
                f"  {d.getName()} ({d.getCount()}); ({'active' if d.isActive() else 'na'}); '{d.getLocale().Country}'; {cls.get_dict_type(d.getDictionaryType())}"  # type: ignore
            )
        print()

    @staticmethod
    def get_dict_type(dt: DictionaryType) -> str:
        """
        Gets dictionary type.

        |lo_safe|

        Args:
            dt (DictionaryType): Dictionary Type

        Returns:
            str: positive, negative, mixed, or ?? if unknown
        """
        if dt == DictionaryType.POSITIVE:
            return "positive"
        if dt == DictionaryType.NEGATIVE:
            return "negative"
        return "mixed" if dt == DictionaryType.MIXED else "??"

    @staticmethod
    def print_con_dicts_info(cd_lst: XConversionDictionaryList) -> None:
        """
        Prints Conversion dictionary list to console.

        |lo_safe|

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
        Gets Lingu Properties.

        |lo_unsafe|

        Raises:
            CreateInstanceMcfError: If unable to create ``com.sun.star.linguistic2.LinguProperties`` instance

        Returns:
            XLinguProperties: Properties
        """
        return mLo.Lo.create_instance_mcf(XLinguProperties, "com.sun.star.linguistic2.LinguProperties", raise_err=True)

    # endregion ---------  Linguistic API ------------------------------

    # region ------------- Linguistics: spell checking -----------------

    @staticmethod
    def load_spell_checker() -> XSpellChecker:
        """
        Gets spell checker.

        |lo_unsafe|

        Raises:
            CreateInstanceMcfError: If unable to create ``com.sun.star.linguistic2.LinguServiceManager`` instance

        Returns:
            XSpellChecker: spell checker
        """
        # lingo_mgr = mLo.Lo.create_instance_mcf(
        #     XLinguServiceManager, "com.sun.star.linguistic2.LinguServiceManager", raise_err=True
        # )
        # return lingo_mgr.getSpellChecker()
        return mLo.Lo.create_instance_mcf(XSpellChecker, "com.sun.star.linguistic2.SpellChecker", raise_err=True)

    # region spell_sentence()

    @overload
    @classmethod
    def spell_sentence(cls, sent: str, speller: XSpellChecker) -> int:
        """
        Spell Check sentence for en US

        Args:
            sent (str): Sentence to spell check
            speller (XSpellChecker): spell checker instance

        Returns:
            int: Number of words spelled incorrectly
        """
        ...

    @overload
    @classmethod
    def spell_sentence(cls, sent: str, speller: XSpellChecker, loc: Locale) -> int:
        """
        Spell Check sentence for en US

        Args:
            sent (str): Sentence to spell check
            speller (XSpellChecker): spell checker instance
            loc (Locale): Local used to spell words.

        Returns:
            int: Number of words spelled incorrectly
        """
        ...

    @classmethod
    def spell_sentence(cls, sent: str, speller: XSpellChecker, loc: Locale | None = None) -> int:
        """
        Spell Check sentence for en US

        Args:
            sent (str): Sentence to spell check
            speller (XSpellChecker): spell checker instance
            loc (Locale | None, optional): Local used to spell words. Default ``Locale("en", "US", "")``

        Returns:
            int: Number of words spelled incorrectly
        """
        # https://tinyurl.com/y6o8doh2
        words = re.split(r"\W+", sent)
        count = 0
        for word in words:
            if loc is None:
                is_correct = cls.spell_word(word=word, speller=speller)
            else:
                is_correct = cls.spell_word(word=word, speller=speller, loc=loc)
            count = count + (0 if is_correct else 1)
        return count

    # endregion spell_sentence()

    # region spell_word()

    @overload
    @staticmethod
    def spell_word(word: str, speller: XSpellChecker) -> bool:
        """
        Spell Check a word for en US.

        |lo_safe|

        Args:
            word (str): word to spell check.
            speller (XSpellChecker): spell checker instance.

        Returns:
            bool: ``True`` if no spelling errors are detected; Otherwise, ``False``.
        """
        ...

    @overload
    @staticmethod
    def spell_word(word: str, speller: XSpellChecker, loc: Locale) -> bool:
        """
        Spell Check a word for en US.

        |lo_safe|

        Args:
            word (str): word to spell check.
            speller (XSpellChecker): spell checker instance.
            loc (Locale | None): Local used to spell word.

        Returns:
            bool: ``True`` if no spelling errors are detected; Otherwise, ``False``.
        """
        ...

    @staticmethod
    def spell_word(word: str, speller: XSpellChecker, loc: Locale | None = None) -> bool:
        """
        Spell Check a word for en US.

        |lo_safe|

        Args:
            word (str): word to spell check.
            speller (XSpellChecker): spell checker instance.
            loc (Locale | None, optional): Local used to spell word. Default ``Locale("en", "US", "")``.

        Returns:
            bool: ``True`` if no spelling errors are detected; Otherwise, ``False``.
        """
        if loc is None:
            loc = Locale("en", "US", "")
        alts = speller.spell(word, loc, ())  # type: ignore
        if alts is not None:
            mLo.Lo.print(f"* '{word}' is unknown. Try:")
            alt_words = alts.getAlternatives()
            mLo.Lo.print_names(alt_words)
            return False
        return True

    # endregion spell_word()

    # endregion ---------- Linguistics: spell checking -----------------

    # region ------------- Linguistics: thesaurus ----------------------

    @staticmethod
    def load_thesaurus() -> XThesaurus:
        """
        Gets Thesaurus.

        |lo_unsafe|

        Raises:
            CreateInstanceMcfError: If unable to create ``com.sun.star.linguistic2.LinguServiceManager`` instance

        Returns:
            XThesaurus: Thesaurus.
        """
        lingo_mgr = mLo.Lo.create_instance_mcf(
            XLinguServiceManager, "com.sun.star.linguistic2.LinguServiceManager", raise_err=True
        )
        return lingo_mgr.getThesaurus()

    @overload
    @staticmethod
    def print_meaning(word: str, thesaurus: XThesaurus) -> int:
        """
        Prints word meanings found in thesaurus to console.

        |lo_safe|

        Args:
            word (str): Word to print meanings of.
            thesaurus (XThesaurus): thesaurus instance.

        Returns:
            int: Number of meanings found.
        """
        ...

    @overload
    @staticmethod
    def print_meaning(word: str, thesaurus: XThesaurus, loc: Locale) -> int:
        """
        Prints word meanings found in thesaurus to console.

        |lo_safe|

        Args:
            word (str): Word to print meanings of.
            thesaurus (XThesaurus): thesaurus instance.
            loc (Locale | None): Local used to query meanings.

        Returns:
            int: Number of meanings found.
        """
        ...

    @staticmethod
    def print_meaning(word: str, thesaurus: XThesaurus, loc: Locale | None = None) -> int:
        """
        Prints word meanings found in thesaurus to console.

        |lo_safe|

        Args:
            word (str): Word to print meanings of.
            thesaurus (XThesaurus): thesaurus instance.
            loc (Locale | None, optional): Local used to query meanings. Default ``Locale("en", "US", "")``.

        Returns:
            int: Number of meanings found.
        """
        if loc is None:
            loc = Locale("en", "US", "")
        meanings = thesaurus.queryMeanings(word, loc, ())  # type: ignore
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
        Gets Proof Reader.

        |lo_unsafe|

        Raises:
            CreateInstanceMcfError: If unable to create linguistic2.Proofreader instance.

        Returns:
            XProofreader: Proof Reader.
        """
        return mLo.Lo.create_instance_mcf(XProofreader, "com.sun.star.linguistic2.Proofreader", raise_err=True)

    @overload
    @classmethod
    def proof_sentence(cls, sent: str, proofreader: XProofreader) -> int:
        """
        Proofs a sentence for en US.

        |lo_safe|

        Args:
            sent (str): sentence to proof.
            proofreader (XProofreader): Proof reader instance.

        Returns:
            int: Number of word of sentence that did not pass proof reading.
        """
        ...

    @overload
    @classmethod
    def proof_sentence(cls, sent: str, proofreader: XProofreader, loc: Locale) -> int:
        """
        Proofs a sentence for en US.

        |lo_safe|

        Args:
            sent (str): sentence to proof.
            proofreader (XProofreader): Proof reader instance.
            loc (Locale | None): Local used to do proof reading.

        Returns:
            int: Number of word of sentence that did not pass proof reading.
        """
        ...

    @classmethod
    def proof_sentence(cls, sent: str, proofreader: XProofreader, loc: Locale | None = None) -> int:
        """
        Proofs a sentence for en US.

        |lo_safe|

        Args:
            sent (str): sentence to proof.
            proofreader (XProofreader): Proof reader instance.
            loc (Locale | None, optional): Local used to do proof reading. Default ``Locale("en", "US", "")``.

        Returns:
            int: Number of word of sentence that did not pass proof reading.
        """
        if loc is None:
            loc = Locale("en", "US", "")
        pr_res = proofreader.doProofreading("1", sent, loc, 0, len(sent), ())
        num_errs = 0
        if pr_res is not None:
            errs = pr_res.aErrors
            if len(errs) > 0:
                for err in errs:
                    cls.print_proof_error(sent, err)
                    num_errs += 1
        return num_errs

    @staticmethod
    def print_proof_error(string: str, err: SingleProofreadingError) -> None:
        """
        Prints proof errors to console.

        |lo_safe|

        Args:
            string (str): error string.
            err (SingleProofreadingError): Single proof reading error.
        """
        e_end = err.nErrorStart + err.nErrorLength
        err_txt = string[err.nErrorStart : e_end]
        print(f"G* {err.aShortComment} in: '{err_txt}'")
        if len(err.aSuggestions) > 0:
            print(f"  Suggested change: '{err.aSuggestions[0]}'")
        print()

    # endregion ---------- Linguistics: grammar checking ---------------

    # region ------------- Linguistics: location guessing --------------

    @staticmethod
    def guess_locale(test_str: str) -> Locale | None:
        """
        Guesses Primary Language and returns results.

        |lo_unsafe|

        Args:
            test_str (str): text used to make guess.

        Returns:
            Locale | None: Local if guess succeeds; Otherwise, ``None``.
        """
        guesser = mLo.Lo.create_instance_mcf(XLanguageGuessing, "com.sun.star.linguistic2.LanguageGuessing")
        if guesser is None:
            mLo.Lo.print("No language guesser found")
            return None
        return guesser.guessPrimaryLanguage(test_str, 0, len(test_str))

    @staticmethod
    def print_locale(loc: Locale) -> None:
        """
        Prints a locale to the console.

        |lo_safe|

        Args:
            loc (Locale): Locale to print.
        """
        if loc is not None:
            print(f"Locale lang: '{loc.Language}'; country: '{loc.Country}'; variant: '{loc.Variant}'")

    # endregion ---------- Linguistics: location guessing --------------

    # region ------------- Linguistics dialogs and menu items ----------

    @staticmethod
    def open_sent_check_options() -> None:
        """
        Open Options - Language Settings - English sentence checking.

        |lo_unsafe|

        Returns:
            None:

        Attention:
            :py:meth:`Lo.dispatch_cmd <.utils.lo.Lo.dispatch_cmd>` method is called along with any of its events.
        """
        pip = mInfo.Info.get_pip()
        lang_ext = pip.getPackageLocation("org.openoffice.en.hunspell.dictionaries")
        mLo.Lo.print(f"Lang Ext: {lang_ext}")
        url = f"{lang_ext}/dialog/en.xdl"
        props = mProps.Props.make_props(OptionsPageURL=url)
        mLo.Lo.dispatch_cmd(cmd="OptionsTreeDialog", props=props)
        mLo.Lo.wait(2000)

    @staticmethod
    def open_spell_grammar_dialog() -> None:
        """
        Activate dialog in  Tools > Spelling and Grammar...

        |lo_unsafe|

        Returns:
            None:

        Attention:
            :py:meth:`Lo.dispatch_cmd <.utils.lo.Lo.dispatch_cmd>` method is called along with any of its events.
        """
        mLo.Lo.dispatch_cmd("SpellingAndGrammarDialog")
        mLo.Lo.wait(2000)

    @staticmethod
    def toggle_auto_spell_check() -> None:
        """
        Toggles spell check on and off.

        |lo_unsafe|

        Returns:
            None:

        Attention:
            :py:meth:`Lo.dispatch_cmd <.utils.lo.Lo.dispatch_cmd>` method is called along with any of its events.
        """
        mLo.Lo.dispatch_cmd("SpellOnline")

    @staticmethod
    def open_thesaurus_dialog() -> None:
        """
        Opens LibreOffice Thesaurus Dialog.

        |lo_unsafe|

        Returns:
            None:

        Attention:
            :py:meth:`Lo.dispatch_cmd <.utils.lo.Lo.dispatch_cmd>` method is called along with any of its events.
        """
        mLo.Lo.dispatch_cmd("ThesaurusDialog")

    # endregion

    @classproperty
    def active_doc(cls) -> XTextDocument:
        """
        Gets current active document.

        |lo_unsafe|

        Returns:
            XTextDocument: Text Document

        .. versionadded:: 0.9.0
        """
        # note:
        # It is not permitted to create weak ref to pyuno objects.
        return mSel.Selection.active_doc


__all__ = ("Write",)
