# coding: utf-8
# Python conversion of Info.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
import os
from typing import cast, overload
from enum import IntEnum


from ..events.event_singleton import _Events
from ..events.lo_named_event import LoNamedEvent
from ..utils import lo as mLo
from ..utils import info as mInfo

from ..meta.static_meta import StaticProperty, classproperty
from ..utils.type_var import DocOrText, DocOrCursor
from ..exceptions import ex as mEx
from ..events.event_singleton import _Events
from ..events.args.event_args import EventArgs
from ..events.args.cancel_event_args import CancelEventArgs
from ..events.write_named_event import WriteNamedEvent


# if not _DOCS_BUILDING and not _ON_RTD:
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XIndexAccess
from com.sun.star.frame import XModel
from com.sun.star.i18n import XBreakIterator
from com.sun.star.text import XParagraphCursor
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XText
from com.sun.star.text import XTextCursor
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.text import XTextRangeCompare
from com.sun.star.text import XTextViewCursor
from com.sun.star.text import XTextViewCursorSupplier
from com.sun.star.text import XWordCursor
from com.sun.star.view import XSelectionSupplier

from ooo.dyn.i18n.word_type import WordTypeEnum as WordTypeEnum
from ooo.dyn.i18n.boundary import Boundary  # struct
from ooo.dyn.lang.locale import Locale  # struct


class Selection(metaclass=StaticProperty):
    """Selection Framework"""

    class CompareEnum(IntEnum):
        """Compare Enumeration"""

        AFTER = 1
        BEFORE = -1
        EQUAL = 0

    @staticmethod
    def is_anything_selected(text_doc: XTextDocument) -> bool:
        """
        Determine if anything is selected.

        If Write document is not visible this method returns false.

        Args:
            text_doc (XTextDocument): Text Document

        Returns:
            bool: True if anything in the document is selected: Otherwise, False

        Note:
            Writer must be visible for this method or ``False`` is always returned.
        """

        model = mLo.Lo.qi(XModel, text_doc, True)

        o_selections = model.getCurrentSelection()
        if o_selections is None:
            return False
        index_access = mLo.Lo.qi(XIndexAccess, o_selections)
        if index_access is None:
            return False
        count = int(index_access.getCount())
        if count == 0:
            return False
        elif count > 1:
            return True
        else:
            # There is only one selection so obtain the first selection
            o_sel = mLo.Lo.qi(XTextRange, index_access.getByIndex(0))
            if o_sel is None:
                return False
            o_text = o_sel.getText()
            # Create a text cursor that covers the range and then see if it is collapsed
            o_cursor = o_text.createTextCursorByRange(o_sel)
            if not o_cursor.isCollapsed():
                return True

        return False

    @staticmethod
    def get_selected_text_range(text_doc: XTextDocument) -> XTextRange | None:
        """
        Gets the text range for current selection

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If unable to obtain required interface.

        Returns:
            XTextRange | None: If no selection is made then None is returned; Otherwise, Text Range.

        Note:
            Writer must be visible for this method or ``None`` is returned.
        """
        model = mLo.Lo.qi(XModel, text_doc, True)

        o_selections: XIndexAccess = model.getCurrentSelection()
        if not o_selections:
            return None
        count = int(o_selections.getCount())
        if count == 0:
            return None
        return mLo.Lo.qi(XTextRange, o_selections.getByIndex(0), True)

    @classmethod
    def get_selected_text_str(cls, text_doc: XTextDocument) -> str:
        """
        Gets the first selection text for Document

        Args:
            text_doc (XTextDocument): Text Document

        Returns:
            str: Selected text or empty string.

        Note:
            Writer must be visible for this method or empty string is returned.
        """
        rng = cls.get_selected_text_range(text_doc=text_doc)
        if rng is None:
            return ""
        return rng.getString()

    @classmethod
    def compare_cursor_ends(cls, c1: XTextRange, c2: XTextRange) -> Selection.CompareEnum:
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
        range_compare = cast(XTextRangeCompare, cls.text_range_compare)
        i = range_compare.compareRegionEnds(c1, c2)
        if i == 1:
            return cls.CompareEnum.BEFORE
        if i == -1:
            return cls.CompareEnum.AFTER
        if i == 0:
            return cls.CompareEnum.EQUAL
        # if no valid result raise error
        msg = "get_cursor_compare_ends() unable to get a valid compare result"
        raise Exception(msg)

    @classmethod
    def range_len(cls, text_doc: XTextDocument, o_sel: XTextRange) -> int:
        """
        Gets the distance between range start and range end.

        Args:
            o_sel (XTextRange): first cursor range
            o_text (object): XText object, usually document text object

        Returns:
            int: length of range

        Note:
            All characters are counted including paragraph breaks.
            In Writer it will display selected characters however,
            paragraph breaks are not counted.
        """
        i = 0
        if o_sel.isCollapsed():
            return i
        o_text = mLo.Lo.qi(XText, text_doc, True)
        l_cursor = cls.get_left_cursor(o_sel=o_sel, o_text=o_text)
        r_cursor = cls.get_right_cursor(o_sel=o_sel, o_text=o_text)
        if cls.compare_cursor_ends(c1=l_cursor, c2=r_cursor) < cls.CompareEnum.EQUAL:

            while cls.compare_cursor_ends(c1=l_cursor, c2=r_cursor) != cls.CompareEnum.EQUAL:
                l_cursor.goRight(1, False)
                i += 1
        return i

    # region ------------- model cursor methods ------------------------

    @classmethod
    def get_text_cursor_props(cls, text_doc: XTextDocument) -> XPropertySet:
        """
        Gets properties for document cursor

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If unable to obtain XPropertySet interface from cursor.

        Returns:
            XPropertySet: Properties
        """
        cursor = cls.get_cursor(text_doc)
        props = mLo.Lo.qi(XPropertySet, cursor, True)
        return props

    # region    get_cursor()
    @overload
    @classmethod
    def get_cursor(cls, cursor_obj: DocOrCursor) -> XTextCursor:
        """
        Gets text cursor

        Args:
            cursor_obj (DocOrCursor): Text Document or Text Cursor

        Returns:
            XTextCursor: Cursor
        """
        ...

    @overload
    @classmethod
    def get_cursor(cls, rng: XTextRange, txt: XText) -> XTextCursor:
        """
        Gets text cursor

        Args:
            rng (XTextRange): Text Range Instance
            txt (XText): Text Instance

        Returns:
            XTextCursor: Cursor
        """
        ...

    @overload
    @classmethod
    def get_cursor(cls, rng: XTextRange, text_doc: XTextDocument) -> XTextCursor:
        """
        Gets text cursor

        Args:
            rng (XTextRange): Text Range instance
            text_doc (XTextDocument): Text Document instance

        Returns:
            XTextCursor: Cursor
        """
        ...

    @classmethod
    def get_cursor(cls, *args, **kwargs) -> XTextCursor:
        """
        Gets text cursor

        Args:
            cursor_obj (DocOrCursor): Text Document or Text View Cursor
            rng (XTextRange): Text Range Instance
            text_doc (XTextDocument): Text Document instance

        Raises:
            CursorError: If Unable to get cursor

        Returns:
            XTextCursor: Cursor
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cursor_obj", "rng", "txt", "text_doc")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_cursor() got an unexpected keyword argument")

            keys = ("cursor_obj", "rng")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("txt", "text_doc")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if not count in (1, 2):
            raise TypeError("get_cursor() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_cursor_obj(kargs[1])
        txt_doc = mLo.Lo.qi(XTextDocument, kargs[2])
        if txt_doc is None:
            txt = kargs[2]
        else:
            txt = txt_doc.getText()
        return cls._get_cursor_txt(rng=kargs[1], txt=txt)

    @staticmethod
    def _get_cursor_txt(rng: XTextRange, txt: XText) -> XTextCursor:
        try:
            cursor = txt.createTextCursorByRange(rng)
            if cursor is None:
                raise Exception("Cursor is None")
            return cursor
        except Exception as e:
            raise mEx.CursorError(str(e)) from e

    @staticmethod
    def _get_cursor_obj(cursor_obj: DocOrCursor) -> XTextCursor:
        try:
            # https://wiki.openoffice.org/wiki/Writer/API/Text_cursor
            xtxt_cursor = mLo.Lo.qi(XTextCursor, cursor_obj)
            if xtxt_cursor is not None:
                c = xtxt_cursor.getText().createTextCursorByRange(xtxt_cursor)
                if c is None:
                    raise Exception("XTextViewCursor.createTextCursorByRange() result is null")
                return c

            xdoc = mLo.Lo.qi(XTextDocument, cursor_obj)
            if xdoc is not None:
                xtext = xdoc.getText()
                if xtext is None:
                    return None
                c = xtext.createTextCursor()
                if c is None:
                    raise Exception("XTextDocument, XText failed to create text cursor")
                return c
            raise TypeError("cursor_obj is invalid type. Expected XTextDocument or XTextViewCursor")
        except Exception as e:
            raise mEx.CursorError(str(e)) from e
        # XTextViewCursor

    # endregion get_cursor()

    @classmethod
    def get_word_cursor(cls, cursor_obj: DocOrCursor) -> XWordCursor:
        """
        Gets document word cursor

        Args:
            cursor_obj (DocOrCursor): Text Document or Text Cursor

        Raises:
            WordCursorError: If Unable to get cursor

        Returns:
            XWordCursor: Word Cursor
        """
        try:
            if mLo.Lo.is_uno_interfaces(cursor_obj, XTextDocument):
                cursor = cls.get_cursor(cursor_obj)
            else:
                cursor = cursor_obj
            wd_cursor = mLo.Lo.qi(XWordCursor, cursor, True)
            return wd_cursor
        except Exception as e:
            raise mEx.WordCursorError(str(e)) from e

    @classmethod
    def get_sentence_cursor(cls, cursor_obj: DocOrCursor) -> XSentenceCursor:
        """
        Gets document sentence cursor

        Args:
            cursor_obj (DocOrCursor): Text Document or Text Cursor

        Raises:
            SentenceCursorError: If Unable to get cursor

        Returns:
            XSentenceCursor: Sentence Cursor
        """
        try:
            if mLo.Lo.is_uno_interfaces(cursor_obj, XTextDocument):
                cursor = cls.get_cursor(cursor_obj)
            else:
                cursor = cursor_obj
            if cursor is None:
                print("Text cursor is null")
                return None
            sc = mLo.Lo.qi(XSentenceCursor, cursor, True)
            return sc
        except Exception as e:
            raise mEx.SentenceCursorError(str(e)) from e

    @classmethod
    def get_paragraph_cursor(cls, cursor_obj: DocOrCursor) -> XParagraphCursor:
        """
        Gets document paragraph cursor

        Args:
            cursor_obj (DocOrCursor): Text Document or Text Cursor

        Raises:
            ParagraphCursorError: If Unable to get cursor

        Returns:
            XParagraphCursor: Paragraph cursor
        """
        try:
            if mLo.Lo.qi(XTextDocument, cursor_obj) is None:
                cursor = cursor_obj
            else:
                cursor = cls.get_cursor(cursor_obj)
            para_cursor = mLo.Lo.qi(XParagraphCursor, cursor, True)
            return para_cursor
        except Exception as e:
            raise mEx.ParagraphCursorError(str(e)) from e

    @classmethod
    def get_left_cursor(cls, o_sel: XTextRange, o_text: DocOrText) -> XTextCursor:
        """
        Creates a new TextCursor with position left that can travel right

        Args:
            o_sel (XTextRange): Text Range
            o_text (DocOrText): Text document or text.

        Returns:
            XTextCursor: a new instance of a TextCursor which is located at the specified
            TextRange to travel in the given text context.
        """
        doc = mLo.Lo.qi(XTextDocument, o_text)
        if doc is None:
            text = o_text
        else:
            text = doc.getText()

        range_compare = cls.text_range_compare
        if range_compare.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
            o_range = o_sel.getEnd()
        else:
            o_range = o_sel.getStart()
        cursor = text.createTextCursorByRange(o_range)
        cursor.goRight(0, False)
        return cursor

    @classmethod
    def get_right_cursor(cls, o_sel: XTextRange, o_text: DocOrText) -> XTextCursor:
        """
        Creates a new TextCursor with position right that can travel left

        Args:
            o_sel (XTextRange): Text Range
            o_text (DocOrText): Text document or text.

        Returns:
            XTextCursor: a new instance of a TextCursor which is located at the specified
            TextRange to travel in the given text context.
        """
        range_compare = cls.text_range_compare
        if range_compare.compareRegionStarts(o_sel.getEnd(), o_sel) >= 0:
            o_range = o_sel.getStart()
        else:
            o_range = o_sel.getEnd()
        cursor = o_text.createTextCursorByRange(o_range)
        cursor.goLeft(0, False)
        return cursor

    @classmethod
    def get_position(cls, cursor: XTextCursor) -> int:
        """
        Gets position of the cursor

        Args:
            cursor (XTextCursor): Text Cursor

        Returns:
            int: Current Cursor Position
        """
        # def get_near_max(l:XTextCursor, r: XTextCursor, jump=10) -> int:
        #     imax = 0
        #     if cls.compare_cursor_ends(l, r) == cls.CompareEnum.BEFORE:
        #         l.goRight(jump, False)
        #         imax = imax + jump
        jmp_amt = 25

        def get_high(l: XTextCursor, r: XTextCursor, jump=jmp_amt, total=0) -> int:
            # OPTIMIZE: get_position.get_high()
            # The idea of this function is to cut down on the number if iterations
            # needed to get the range from cursors left and right positions.
            # Most likely there is an even more efficient way to do this.
            if jump <= 0:
                return 0
            if cls.compare_cursor_ends(l, r) == cls.CompareEnum.BEFORE:
                j = jump + jmp_amt
                l.goRight(j, False)
                return get_high(l, r, j, jump)
            else:
                l.gotoStart(False)
                l.goRight(total, False)
                return total
            # else:
            #     return jump, True

        xcursor = cursor.getText().createTextCursor()
        xcursor.gotoStart(False)
        xcursor.gotoEnd(True)
        xtext = xcursor.getText()
        l_cursor = cls.get_left_cursor(o_sel=xcursor, o_text=xtext)
        r_cursor = cls.get_right_cursor(o_sel=xcursor, o_text=xtext)
        i = 0
        if cls.compare_cursor_ends(c1=l_cursor, c2=r_cursor) < cls.CompareEnum.EQUAL:
            high = get_high(l_cursor, r_cursor)
            # l_cursor.gotoStart(False)

            while cls.compare_cursor_ends(c1=l_cursor, c2=r_cursor) != cls.CompareEnum.EQUAL:
                l_cursor.goRight(1, False)
                i += 1
            i += high
        return i

        # return len(cursor.getText().getString())

    # region ------------- view cursor methods -------------------------

    @classmethod
    def get_text_view_cursor_prop_set(cls, text_doc: XTextDocument) -> XPropertySet:
        """
        Gets properties for document view cursor

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            MissingInterfaceError: If unable to obtain XPropertySet interface from cursor.

        Returns:
            XPropertySet: Properties
        """
        xview_cursor = cls.get_view_cursor(text_doc)
        props = mLo.Lo.qi(XPropertySet, xview_cursor, True)
        return props

    @staticmethod
    def get_view_cursor(text_doc: XTextDocument) -> XTextViewCursor:
        """
        Gets document view cursor.

        Describes a cursor in a text document's view.

        Args:
            text_doc (XTextDocument): Text Document

        Raises:
            ViewCursorError: If Unable to get cursor

        Returns:
            XTextViewCursor: Text View Cursor

        See Also:
            `LibreOffice API XTextViewCursor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextViewCursor.html>`_
        """

        # https://wiki.openoffice.org/wiki/Writer/API/Text_cursor
        try:
            model = mLo.Lo.qi(XModel, text_doc, True)
            xcontroller = model.getCurrentController()
            supplier = mLo.Lo.qi(XTextViewCursorSupplier, xcontroller, True)
            vc = supplier.getViewCursor()
            if vc is None:
                raise Exception("Supplier return null view cursor")
            return vc
        except Exception as e:
            raise mEx.ViewCursorError(str(e)) from e

    # endregion ---------- view cursor methods -------------------------

    # endregion ---------- model cursor methods ------------------------

    @classproperty
    def text_range_compare(cls) -> XTextRangeCompare:
        """
        Gets text range for comparison operations

        Returns:
            XTextRangeCompare: Text Range Compare instance
        """

        try:
            return cls._text_range_compare
        except AttributeError:
            doc = mLo.Lo.XSCRIPTCONTEXT.getDocument()
            text = doc.getText()
            cls._text_range_compare = mLo.Lo.qi(XTextRangeCompare, text)
        return cls._text_range_compare

    @text_range_compare.setter
    def text_range_compare(cls, value) -> None:
        # raise error on set. Not really necessary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.__name__)

    @staticmethod
    def get_word_count_ooo(text: str, word_type: WordTypeEnum | None = None, locale_lang: str | None = None) -> int:
        """
        Get the number of word in ooo way.

        This method takes into account the current Locale

        Args:
            text (str): string to count the word of
            word_type (WordTypeEnum, optional): type of words to count. Default ``WordTypeEnum.WORD_COUNT``
                Import  line ``from ooodev.utils.selection import WordTypeEnum``
            locale_lang (str, optional): Language such as 'en-US' used to process word boundaries. Defaults to LO's current language.

        Raises:
            CreateInstanceMsfError: If unable to create ``i18n.BreakIterator service``

        Returns:
            int: The number of words
        """
        if word_type is None:
            word_type = WordTypeEnum.WORD_COUNT
        # https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=82678
        next_wd = Boundary()
        local = Locale()
        # local.Language = "en"
        if locale_lang is None:
            local.Language = mInfo.Info.language
        else:
            local.Language = locale_lang
        num_words = 0
        start_pos = 0
        if word_type > WordTypeEnum.ANY_WORD:
            # intensionally pad the start and end of the string to guarantee we get the first word and clean break on last word
            st = f" {text} "
        else:
            # ANY_WORD
            # no need to pad. All characters will be counted including whitespaces
            st = text

        brk = mLo.Lo.create_instance_mcf(XBreakIterator, "com.sun.star.i18n.BreakIterator")
        if brk is None:
            raise mEx.CreateInstanceMcfError(XBreakIterator, "com.sun.star.i18n.BreakIterator")

        next_wd = brk.nextWord(st, start_pos, local, word_type.value)
        while next_wd.startPos != next_wd.endPos:
            num_words += 1
            nw = next_wd.startPos
            next_wd = brk.nextWord(st, nw, local, word_type.value)

        return num_words

    @classmethod
    def select_next_word(cls, text_doc: XTextDocument) -> bool:
        """
        Select the word right from the current cursor position.

        Args:
            text_doc (XTextDocument): Text Document

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

        cargs = CancelEventArgs(Selection.select_next_word)
        cargs.event_data = {"text_doc": text_doc}
        _Events().trigger(WriteNamedEvent.WORD_SELECTING, cargs)
        if cargs.cancel:
            return False

        supplier = mLo.Lo.qi(XSelectionSupplier, text_doc.getCurrentController(), True)

        # clear any current selection
        view_curor = cls.get_view_cursor(text_doc=text_doc)
        # view_curor.collapseToEnd()
        view_curor.goRight(0, False)

        # see section 7.5.1 of developers' guide
        index_access = mLo.Lo.qi(XIndexAccess, supplier.getSelection(), True)
        range = mLo.Lo.qi(XTextRange, index_access.getByIndex(0), True)

        # get the XWordCursor and make a selection!
        xText = range.getText()
        word_cursor = cls.get_word_cursor(xText.createTextCursorByRange(range))

        if not word_cursor.isStartOfWord():
            word_cursor.gotoStartOfWord(False)

        result = word_cursor.gotoNextWord(True)
        supplier.select(word_cursor)
        _Events().trigger(WriteNamedEvent.WORD_SELECTED, EventArgs.from_args(cargs))
        return result


def _del_cache_attrs(source: object, e: EventArgs) -> None:
    # clears Write Attributes that are dynamically created
    dattrs = ("_text_range_compare",)
    for attr in dattrs:
        if hasattr(Selection, attr):
            delattr(Selection, attr)


# subscribe to events that warrant clearing cached attribs
_Events().on(LoNamedEvent.RESET, _del_cache_attrs)

__all__ = ("Selection",)
