# coding: utf-8
from __future__ import annotations
import sys
from typing import TYPE_CHECKING, Iterable, List, overload, cast
import uno
import re

from ..utils.gen_util import TableHelper
from ..utils import lo as m_lo
from ..utils import info as m_info
from ..utils import file_io as m_file_io
from ..utils import props as m_props
from ..utils import images as m_images

from com.sun.star.awt import FontWeight # type: ignore
from com.sun.star.awt.FontSlant import ITALIC as FS_ITALIC # type: ignore
from com.sun.star.awt import Size # struct
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XEnumerationAccess
from com.sun.star.document import XDocumentInsertable
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XShape
from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import Locale # struct class
from com.sun.star.linguistic2.DictionaryType import POSITIVE as DT_POSITIVE, NEGATIVE as DT_NEGATIVE, MIXED as DT_MIXED # type: ignore enum values
from com.sun.star.style import NumberingType # type: ignore
from com.sun.star.style.BreakType import PAGE_AFTER as BT_PAGE_AFTER, COLUMN_AFTER as BT_COLUMN_AFTER # type: ignore
from com.sun.star.style.ParagraphAdjust import CENTER as PA_CENTER, RIGHT as PA_RIGHT # type: ignore
from com.sun.star.table import BorderLine # struct
from com.sun.star.text import ControlCharacter # type: ignore
from com.sun.star.text import HoriOrientation #type: ignore Const
from com.sun.star.text import VertOrientation #type: ignore const
from com.sun.star.text import XBookmarksSupplier
from com.sun.star.text import XPageCursor
from com.sun.star.text import XParagraphCursor
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XText
from com.sun.star.text import XTextContent
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextGraphicObjectsSupplier
from com.sun.star.text import XTextViewCursor
from com.sun.star.text import XWordCursor
from com.sun.star.text.PageNumberType import CURRENT as PN_CURRENT # type: ignore
from com.sun.star.view.PaperFormat import A4 as PF_A4 # type: ignore
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER # type: ignore
from com.sun.star.text.TextContentAnchorType import AT_PAGE #type: ignore
from com.sun.star.uno import Exception as UnoException

if TYPE_CHECKING:
    from com.sun.star.container import XEnumeration
    from com.sun.star.container import XNameAccess
    from com.sun.star.container import XNamed
    from com.sun.star.document import XEmbeddedObjectSupplier2
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.graphic import XGraphic
    from com.sun.star.text import XTextFrame
    from com.sun.star.frame import XModel
    from com.sun.star.lang import XComponent
    from com.sun.star.linguistic2 import SingleProofreadingError
    from com.sun.star.linguistic2 import XConversionDictionaryList
    from com.sun.star.linguistic2 import XLanguageGuessing
    from com.sun.star.linguistic2 import XLinguProperties
    from com.sun.star.linguistic2 import XLinguServiceManager
    from com.sun.star.linguistic2 import XLinguServiceManager2
    from com.sun.star.linguistic2 import XProofreader
    from com.sun.star.linguistic2 import XSearchableDictionaryList
    from com.sun.star.linguistic2 import XSpellChecker
    from com.sun.star.linguistic2 import XThesaurus
    from com.sun.star.text import XSimpleText
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextField
    from com.sun.star.text import XTextTable
    from com.sun.star.text import XTextViewCursorSupplier
    from com.sun.star.view import XPrintable

if sys.version_info >= (3, 10):
    from typing import Union
else:
    from typing_extensions import Union

Lo = m_lo.Lo
Info = m_info.Info
FileIO = m_file_io.FileIO
Props = m_props.Props
Images = m_images.Images

class Write:
    
    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader) -> XTextDocument | None:
        doc = Lo.open_doc(fnm=fnm, loader=loader)
        if doc is None:
            print("Document is null")
            return None
        if not cls.is_text(doc):
            print(f"Not a text document; closing '{fnm}'")
            return None
        
        if not Info.support_service(doc, XTextDocument):
            print(f"Not a text document; closing '{fnm}'")
            Lo.close_doc(doc)
            return None
        return doc
    
    @staticmethod
    def is_text(doc: XComponent) -> bool:
        return Info.is_doc_type(obj=doc, doc_type=Lo.WRITER_SERVICE)
    
    @staticmethod
    def get_text_doc(doc: XComponent) -> XTextDocument:
        if doc is None:
            print("Document is null")
            return None
        
        if not Info.support_service(doc, XTextDocument):
            print("Not a text document")
            return None
        return doc
    
    @staticmethod
    def create_doc(loader: XComponentLoader) -> XTextDocument:
        return Lo.create_doc(doc_type=Lo.WRITER_STR, loader=loader)
    
    @staticmethod
    def create_doc_from_template(template_path: str, loader: XComponentLoader) -> XTextDocument:
        return Lo.create_doc_from_template(template_path=template_path, loader=loader)
    
    @staticmethod
    def close_doc(text_doc: XTextDocument) -> None:
        Lo.close(text_doc)
    
    @staticmethod
    def save_doc(text_doc: XTextDocument, fnm: str) -> None:
        Lo.save_doc(doc=text_doc, fnm=fnm)
    
    @classmethod
    def open_flat_doc_using_text_template(cls, fnm: str, template_path: str, loader: XComponentLoader) -> XTextDocument | None:
        if fnm is None:
            print("Filename is null")
            return None
        
        open_file_url = None
        if not FileIO.is_openable(fnm):
            if Lo.is_url(fnm):
                print(f"Will treat filename as a URL: '{fnm}'")
                open_file_url = fnm
            else:
                return None
        else:
            open_file_url = FileIO.fnm_to_url(fnm)
            if open_file_url is None:
                return None
        
        template_ext = Info.get_ext(template_path)
        if template_ext != 'ott':
            print("Can only apply a text template as formatting")
            return None

        doc = Lo.create_doc_from_template(template_path==template_path, loader=loader)
        if doc is None:
            return None
        cursor = cast(XTextCursor | XDocumentInsertable, cls.get_cursor(doc))
        try:
            cursor.gotoEnd(True)
            if not Info.support_service(cursor, XDocumentInsertable):
                print("Document inserter could not be created")
            else:
                cursor.insertDocumentFromURL(open_file_url, tuple())
        except Exception as e:
            print("Could not insert document")
            print(f"    {e}")
        return doc
    
    # --------------------- model cursor methods -------------------------------

    @overload
    @staticmethod
    def get_cursor(cursor_obj: XTextDocument) -> XTextCursor | None:...

    @overload
    @staticmethod
    def get_cursor(cursor_obj: XTextViewCursor) -> XTextCursor | None:...
    
    
    @staticmethod
    def get_cursor(cursor_obj: XTextDocument | XTextViewCursor) -> XTextCursor | None:
        xtext = cursor_obj.getText()
        if xtext is None:
            print("Text not found in document")
            return None
        if Info.support_service(cursor_obj, XTextViewCursor):
            return xtext.createTextCursorByRange(cursor_obj)
        return xtext.createTextCursor()
    
    @classmethod
    def get_word_cursor(cls, text_doc: XTextDocument) -> XWordCursor | None:
        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            print("Text cursor is null")
            return None
        if not Info.support_service(cursor, XWordCursor):
            print("Text is not XWordCursor")
            return None
        return cursor
    
    @classmethod
    def get_sentence_cursor(cls, text_doc: XTextDocument) -> XSentenceCursor | None:
        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            print("Text cursor is null")
            return None
        if not Info.support_service(cursor, XSentenceCursor):
            print("Text is not XSentenceCursor")
            return None
        return cursor
    
    @classmethod
    def get_paragraph_cursor(cls, text_doc: XTextDocument) -> XParagraphCursor | None:
        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            print("Text cursor is null")
            return None
        if not Info.support_service(cursor, XParagraphCursor):
            print("Text is not XParagraphCursor")
            return None
        return cursor
    
    @staticmethod
    def get_position(cursor: XTextCursor) -> int:
        return len(cursor.getText().getString())
    
    # ---------------------- view cursor methods ------------------------------
    @staticmethod
    def get_view_cursor(text_doc: XTextDocument | XModel) -> XTextViewCursor:
        xcontroller: XTextViewCursorSupplier = text_doc.getCurrentController()
        return xcontroller.getViewCursor()
    
    @staticmethod
    def get_current_page(tv_cursor: XTextViewCursor | XPageCursor) -> int:
        if not Info.support_service(tv_cursor, XPageCursor):
            print("Could not create a page cursor")
            return -1
        return tv_cursor.getPage()
    
    @staticmethod
    def get_coord_str(tv_cursor: XTextViewCursor) -> str:
        pos = tv_cursor.getPosition()
        return f"({pos.X}, {pos.Y})"
    
    @staticmethod
    def get_page_count(text_doc: XTextDocument | XModel) -> int:
        xcontroller = text_doc.getCurrentController()
        return int(Props.get_property(xcontroller, "PageCount"))
    
    @classmethod
    def get_text_view_cursor_prop_set(cls, text_doc: XTextDocument) -> XPropertySet:
        xview_cursor = cls.get_view_cursor(text_doc)
        return xview_cursor

    # ------------------------- text writing methods ------------------------------------
    
    @classmethod
    def _append1(cls, cursor: XTextCursor, text:str) -> int:
        cursor.setString(text)
        cursor.gotoEnd(False)
        return cls.get_position(cursor)
    
    @classmethod
    def _append2(cls, cursor: XTextCursor, ctl_char: int) -> int:
        xtext = cursor.getText()
        xtext.insertControlCharacter(cursor, ctl_char, False)
        cursor.gotoEnd(False)
        return cls.get_position(cursor)
    
    @classmethod
    def _append3(cls, cursor: XTextCursor, text_content: XTextContent) -> int:
        xtext = cursor.getText()
        xtext.insertTextContent(cursor, text_content, False)
        cursor.gotoEnd(False)
        return cls.get_position(cursor)
    
    @overload
    @staticmethod
    def append(cursor: XTextCursor, text:str) -> int:...
    
    @overload
    @staticmethod
    def append(cursor: XTextCursor, ctl_char:int) -> int:...
    
    @overload
    @staticmethod
    def append(cursor: XTextCursor, text_content: XTextContent) -> int:...
    
    @classmethod
    def append(cls, *args, **kwargs) -> int:
        ordered_keys = ("first", "second")
        kargs = {}
        kargs["first"] = kwargs.get("cursor", None)
        i = 0
        if 'text' in kwargs:
            kargs['second'] = kwargs['second']
            i = 1
        elif 'ctl_char' in kwargs:
            kargs['second'] = kargs['ctl_char']
            i =2
        elif 'text_content' in kwargs:
            kargs["second"] = kwargs['text_content']
            i = 3
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if i == 0:
            # str, int or, XTextContent
            sec = kargs['second']
            if isinstance(sec, str):
                i = 1
            elif isinstance(sec, int):
                i = 2
            else:
                i = 3
        if i == 1:
            return cls._append1(kargs["first"], kargs["second"])
        if i == 2:
            return cls._append2(kargs["first"], kargs["second"])
        return cls._append3(kargs["first"], kargs["second"])

    @classmethod
    def append_date_time(cls, cursor: XTextCursor) -> int:
        """append two DateTime fields, one for the date, one for the time"""
        dt_field: XTextField = Lo.create_instance_mcf("com.sun.star.text.TextField.DateTime")
        Props.set_property(dt_field, "IsDate", True) # so date is reported
        cls.append(cursor, dt_field)
        cls.append(cursor, "; ")
        
        dt_field: XTextField = Lo.create_instance_mcf("com.sun.star.text.TextField.DateTime")
        Props.set_property(dt_field, "IsDate", False) # so time is reported
        return cls.append(cursor, dt_field)
    
    @classmethod
    def append_para(cls, cursor: XTextCursor, text: str) -> int:
        cls.append(cursor=cursor, text=text)
        cls.append(cursor=cursor, ctl_char=ControlCharacter.PARAGRAPH_BREAK)
        return cls.get_position(cursor)
    
    @classmethod
    def end_line(cls, cursor: XTextCursor) -> None:
        cls.append(cursor=cursor, ctl_char=ControlCharacter.LINE_BREAK)
    
    @classmethod
    def end_paragraph(cls, cursor: XTextCursor) -> None:
        cls.append(cursor=cursor, ctl_char=ControlCharacter.PARAGRAPH_BREAK)
    
    @classmethod
    def page_break(cls, cursor: XTextCursor) -> None:
        Props.set_property(cursor, "BreakType", BT_PAGE_AFTER)
        cls.end_paragraph(cursor)
    
    @classmethod
    def column_break(cls, cursor: XTextCursor) -> None:
        Props.set_property(cursor, "BreakType", BT_COLUMN_AFTER)
        cls.end_paragraph(cursor)
    
    @classmethod
    def insert_para(cls, cursor: XTextCursor, para: str, para_style: str) -> None:
        xtext = cursor.getText()
        xtext.insertString(cursor,para, False)
        xtext.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        cls.style_prev_paragraph(cursor, para_style)
    
    # --------------------- extract text from document ------------------------
    
    @staticmethod
    def get_all_text(cursor: XTextCursor) -> str:
        """return the text part of the document"""
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        text = cursor.getString()
        cursor.gotoEnd(False) # to deselect everything
        return text
    
    @staticmethod
    def get_enumeration(obj: XEnumerationAccess) -> XEnumeration | None:
        if not Info.support_service(obj, XEnumerationAccess):
            print("XEnumerationAccess not supported for obj")
            return None
        return obj.createEnumeration()
    
    # ------------------------ text cursor property methods -----------------------------------
    
    @classmethod
    def style_left_bold(cls, cursor: XTextCursor, pos: int) -> None:
        cls.style_left(cursor, pos, "CharWeight", FontWeight.BOLD)
    
    @classmethod
    def style_left_italic(cls, cursor: XTextCursor, pos: int) -> None:
        cls.style_left(cursor, pos, "CharPosture", FS_ITALIC)
    
    @classmethod
    def style_left_color(cls, cursor: XTextCursor, pos: int, color: int) -> None:
        cls.style_left(cursor, pos, "CharColor", color)
    
    @classmethod
    def style_left_code(cls, cursor: XTextCursor, pos: int) -> None:
        cls.style_left(cursor, pos, "CharFontName", Info.get_font_mono_name())
        cls.style_left(cursor, pos, "CharHeight", 10)
    
    @classmethod
    def style_left(cls, cursor: XTextCursor, pos: int, prop_name: str, prop_val: object) -> None:
        old_val = Props.get_property(cursor, prop_name)
        
        curr_pos = cls.get_position(cursor)
        cursor.goLeft(curr_pos - pos, False)
        Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)
        
        cursor.goRight(curr_pos - pos, False)
        Props.set_property(prop_set=cursor, name=prop_name, value=old_val)
    
    @overload
    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor, prop_val: object) -> None:...
    @overload
    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor, prop_val: object, prop_name: str) -> None:...
    
    @staticmethod
    def style_prev_paragraph(cursor: XTextCursor | XParagraphCursor, prop_val: object, prop_name: str = None) -> None:
        if prop_name is None:
            prop_name = "ParaStyleName"
        old_val = Props.get_property(cursor, prop_name)
        
        cursor.gotoPreviousParagraph(True) # select previous paragraph
        Props.set_property(prop_set=cursor, name=prop_name, value=prop_val)
        
        # reset
        cursor.gotoNextParagraph(False)
        Props.set_property(prop_set=cursor, name=prop_name, value=old_val)

    # ---------------------------- style methods -------------------------------
    
    @staticmethod
    def get_page_text_width(text_doc: XTextDocument) -> int:
        """get the width of the page's text area"""
        props = Info.get_style_props(doc=text_doc, family_style_name="PageStyles",prop_set_nm="Standard")
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
    def get_page_size(text_doc: XTextDocument) -> Size | None:
        props = Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            print("Could not access the standard page style")
            return None
        try:
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            return Size(width, height)
        except Exception as e:
            print("Could not access standard page style dimensions")
            print(f"    {e}")
            return None
    
    @staticmethod
    def set_a4_page_format(text_doc: XTextDocument | XPrintable) -> None:
        printer_desc = Props.make_props(PaperFormat=PF_A4)
        
        # Paper Orientation           
        # java
        # printerDesc[1] = new PropertyValue();
        # printerDesc[1].Name = "PaperOrientation";
        # printerDesc[1].Value = PaperOrientation.LANDSCAPE;
        
        text_doc.setPrinter(printer_desc)
    
    # ------------------------ headers and footers ----------------------------
    @classmethod
    def set_page_numbers(cls, text_doc: XTextDocument) -> None:
        """
        Modify the footer via the page style for the document. 
        Put page number & count in the center of the footer in Times New Roman, 12pt
        """
        props = Info.get_style_props(doc=text_doc, family_style_name="PageStyles",prop_set_nm="Standard")
        if props is None:
            print("Could not access the standard page style")
            return None
        
        try:
            props.setPropertyValue("FooterIsOn", True)
                #   footer must be turned on in the document
            footer_text: XText = props.getPropertyValue("FooterText")
            footer_cursor = footer_text.createTextCursor()
            
            Props.set_property(prop_set=footer_cursor, name="CharFontName", value=Info.get_font_general_name())
            Props.set_property(prop_set=footer_cursor, name="CharHeight", value=12.0)
            Props.set_property(prop_set=footer_cursor, name="ParaAdjust", value=PA_CENTER)
        
            # add text fields to the footer
            cls._append3(cursor=footer_cursor,text_content=cls.get_page_number())
            cls._append1(cursor=footer_cursor, text=" of ")
            cls._append3(cursor=footer_cursor,text_content=cls.get_page_count())
        except Exception as e:
            print(e)
            
    @staticmethod
    def get_page_number() -> XTextField:
        """return arabic style number showing current page value"""
        num_field: XTextField = Lo.create_instance_msf("com.sun.star.text.TextField.PageNumber")
        Props.set_property(prop_set=num_field, name="NumberingType", value=NumberingType.ARABIC)
        Props.set_property(prop_set=num_field, name="SubType", value=PN_CURRENT)
        return num_field
    
    @staticmethod
    def get_page_count() -> XTextField:
        """return arabic style number showing current page count"""
        pc_field: XTextField = Lo.create_instance_msf("com.sun.star.text.TextField.PageCount")
        Props.set_property(prop_set=pc_field, name="NumberingType", value=NumberingType.ARABIC)
        return pc_field
    
    @staticmethod
    def set_header(text_doc: XTextDocument, h_text: str) -> None:
        """
        Modify the header via the page style for the document. 
        Put the text on the right hand side in the header in
        a general font of 10pt.
        """
        props = Info.get_style_props(doc=text_doc,family_style_name="PageStyles", prop_set_nm="Standard")
        if props is None:
            print("Could not access the standard page style container")
            return
        try:
            props.setPropertyValue("HeaderIsOn", True)
                    # header must be turned on in the document
            # props.setPropertyValue("TopMargin", 2200)
            header_text: XText = props.getPropertyValue("HeaderText")
            header_cursor = cast(XTextCursor | XPropertySet, header_text.createTextCursor())
            header_cursor.gotoEnd(False)
            
            header_cursor.setPropertyValue("CharFontName", Info.get_font_general_name())
            header_cursor.setPropertyValue("CharHeight", 10)
            header_cursor.setPropertyValue("ParaAdjust", PA_RIGHT)
            
            header_text.setString(f"{h_text}\n")
        except Exception as e:
            print(e)

    @staticmethod
    def get_draw_page(doc: XTextDocument | XDrawPageSupplier) -> XDrawPage | None:
        if not Info.support_service(obj=doc, service=XDrawPageSupplier):
            print("XDrawPageSupplier interface is not supporte for doc")
            return None
        return doc.getDrawPage()

    # -------------------------- adding elements ----------------------------
    
    @classmethod
    def add_formula(cls, cursor: XTextCursor, formula: str) -> None:
        try:
            embed_content: XTextContent = Lo.create_instance_msf("com.sun.star.text.TextEmbeddedObject")
            if embed_content is None:
                print("Could not create a formula embedded object")
                return
            # set class ID for type of object being inserted
            props = cast(XPropertySet, embed_content)
            props.setPropertyValue("CLSID", Lo.MATH_CLSID)
            props.setPropertyValue("AnchorType", AS_CHARACTER)
            
            # insert object in document
            cls._append3(cursor=cursor, text_content=embed_content)
            cls.end_line(cursor)
            
            # access object's model
            embed_obj_supplier = cast(XEmbeddedObjectSupplier2, embed_content)
            embed_obj_model = embed_obj_supplier.getEmbeddedObject()
            formula_props = cast(XPropertySet, embed_obj_model)
            formula_props.setPropertyValue("Formula", formula)
            print(f'Inserted formula "{formula}"')
        except Exception as e:
            print(f'Insertion fo formula "{formula}" failed:')
            print(f"    {e}")
    
    @classmethod
    def add_hyperlink(cls, cursor: XTextCursor, label: str, url_str: str) -> None:
        link: XTextContent = Lo.create_instance_msf("com.sun.star.text.TextField.URL")
        if link is None:
            print("Could not create a hyperlink")
            return
        Props.set_property(prop_set=link, name= "URL", value=url_str)
        Props.set_property(prop_set=link, name= "Representation", value=label)
        
        cls._append3(cursor, link)
    
    @classmethod
    def add_bookmark(cls, cursor: XTextCursor, name: str) -> None:
        bmk_content: XTextContent = Lo.create_instance_msf("com.sun.star.text.Bookmark")
        if bmk_content is None:
            print("Could not create a bookmark")
            return
        
        bmk_named = cast(XNamed, bmk_content)
        bmk_named.setName(name)
        
        cls._append3(cursor, bmk_content)
    
    @staticmethod
    def find_bookmark(doc: XTextDocument, bm_name:str) -> XTextContent | None:
        supplier = Lo.qi(XBookmarksSupplier, doc)
        
        if supplier is None:
            print("Bookmark supplier could not be created")
            return None
        
        named_bookmarks = supplier.getBookmarks()
        if named_bookmarks is None:
            print("Name access to bookmarks not possible")
            return None
        
        obookmark = None
        try:
            obookmark = named_bookmarks.getByName(bm_name)
        except Exception:
            print(f"Bookmark '{bm_name}' not found")
            return None
        return Lo.qi(XTextContent, obookmark)
    
    @classmethod
    def add_text_frame(cls, cursor: XTextCursor, ypos: int, text: str, width: int, height: int) -> None:
        try:
            xframe: XTextFrame = Lo.create_instance_msf("com.sun.star.text.TextFrame")
            tf_shape = Lo.qi(XShape, xframe)
            if tf_shape is None:
                return
            
            tf_shape.setSize(Size(width, height))
            frame_props = Lo.qi(XPropertySet, xframe)
            if frame_props is None:
                return
            frame_props.setPropertyValue("AnchorType", AT_PAGE)
            frame_props.setPropertyValue("FrameIsAutomaticHeight", True) # will grow if necessary
            
            border = BorderLine()
            border.OuterLineWidth = 1
            border.Color = 0xFF0000;  # red
            
            frame_props.setPropertyValue("TopBorder", border)
            frame_props.setPropertyValue("BottomBorder", border)
            frame_props.setPropertyValue("LeftBorder", border)
            frame_props.setPropertyValue("RightBorder", border)
            
            # make the text frame blue
            frame_props.setPropertyValue("BackTransparent", False) # not transparent
            frame_props.setPropertyValue("RightBorder", 0xCCCCFF) # light blue
            
            # Set the horizontal and vertical position
            frame_props.setPropertyValue("HoriOrient", HoriOrientation.RIGHT)
            frame_props.setPropertyValue("VertOrient", VertOrientation.NONE)
            frame_props.setPropertyValue("VertOrientPosition", ypos)
            
            cls._append3(cursor, xframe)
            cls.end_paragraph(cursor)
            
            # add text into the text frame
            xframe_text = xframe.getText()
            xframe_cursor = cast(XSimpleText, xframe_text.createTextCursor())
            xframe_cursor.insertString(xframe_cursor, text, False)
        except Exception as e:
            print("Insertion of text frame failed:")
            print(f"    {e}")
    
    @classmethod
    def add_table(cls, cursor: XTextCursor, rows_list: List[List[str]]) -> None:
        """
        Each row becomes a row of the table. The first row is treated as a header,
        and colored in dark blue, and the rest in light blue. 
        """
        try:
            table: XTextTable = Lo.create_instance_msf("com.sun.star.text.TextTable")
            if table is None:
                print("Could not create a text table")
                return
            num_rows = len(rows_list)
            if num_rows == 0:
                return
            num_cols = len(rows_list[0])
            print(f"Creating table rows: {num_rows}, cols: {num_cols}")
            table.initialize(num_rows, num_cols)
            
            # insert the table into the document
            cls._append3(cursor, table)
            cls.end_paragraph(cursor)
            
            table_props = Lo.qi(XPropertySet, table)
            if table_props is None:
                return None
            
            # set table properties
            table_props.setPropertyValue("BackTransparent", False) # not transparent
            table_props.setPropertyValue("BackColor", 0xCCCCFF) # light blue
            
            # set color of first row (i.e. the header) to be dark blue
            rows = table.getRows()
            Props.set_property(prop_set=rows.getByIndex(0), name="BackColor", value=0x666694) # dark blue
            
            #  write table header
            row_data = rows_list[0]
            for x in range(num_cols):
                cls.set_cell_header(cls.mk_cell_name(1, x+1), row_data[x], table)
                                    # e.g. "A1", "B1", "C1", etc
            
            # insert table body
            for y in range(1, num_rows): # start in 2nd row
                row_data = rows_list[y]
                for x in range(num_cols):
                    cls.set_cell_text( cls.mk_cell_name(x,y+1), row_data[x], table)
        except Exception as e:
            print("Table insertion failed:")
            print(f"    {e}")
    
    
    @staticmethod
    def mk_cell_name(row: int, col: int) -> str:
        """
        Convert given row and column number to ``A1`` style cell name.

        Args:
            row (int): Row number. This is a 1 based value.
            col (int): Column Number. This is 1 based value.

        Raises:
            ValueError: If row or col value < 1

        Returns:
            str: row and col as cell name such as A1, AB3
        """
        return TableHelper.make_cell_name(row=row, col=col)
    
    @staticmethod
    def set_cell_header(cell_name: str, data: str, table: XTextTable) -> None:
        cell_text = Lo.qi(XText, table.getCellByName(cell_name))
        if cell_text is None:
            return
        text_cursor = cell_text.createTextCursor()
        Props.set_property(prop_set=text_cursor, name="CharColor",value=0xFFFFFF) # use white text
        
        cell_text.setString(str(data))
        
    @staticmethod
    def set_cell_text(cell_name: str, data: str, table: XTextTable) -> None:
        cell_text = Lo.qi(XText, table.getCellByName(cell_name))
        if cell_text is None:
            return
        cell_text.setString(str(data))

    @overload
    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: str) -> None:...
    
    @overload
    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: str, width: int, height: int) -> None:...
    

    @classmethod
    def add_image_link(cls, doc: XTextDocument, cursor: XTextCursor, fnm: str, width: int = 0, height: int = 0) -> None:
        try:
            tgo: XTextContent = Lo.create_instance_msf("com.sun.star.text.TextGraphicObject")
            if tgo is None:
                print("Could not create a text graphic object")
                return
            
            props = Lo.qi(XPropertySet, tgo)
            if props is None:
                print("Could not create a text graphic object")
                return
            props.setPropertyValue("AnchorType", AS_CHARACTER)
            props.setPropertyValue("GraphicURL", FileIO.fnm_to_url(fnm))
            
            # optionally set the width and height
            if width > 0 and height > 0:
                props.setPropertyValue("Width", width)
                props.setPropertyValue("Height", height)
            
            # append image to document, followed by a newline
            cls._append3(cursor, tgo)
            cls.end_line(cursor)
        except Exception as e:
            print(f"Insertion of graphic in '{fnm}' failed:")
            print(f"    {e}")
    
    @overload
    @staticmethod
    def add_image_shape(doc: XTextDocument, cursor: XTextCursor, fnm: str) -> None:...
    
    @overload
    @staticmethod
    def add_image_shape(doc: XTextDocument, cursor: XTextCursor, fnm: str, width: int, height: int) -> None:...
    

    @classmethod
    def add_image_shape(cls, doc: XTextDocument, cursor: XTextCursor, fnm: str, width: int = 0, height: int = 0) -> None:
        if width > 0 and height > 0:
            im_size = Size(width, height)
        else:
            im_size = Images.get_size_100mm(fnm) # in 1/100 mm units
            if im_size is None:
                return
        
        try:
            # create TextContent for an empty graphic
            gos: XTextContent = Lo.create_instance_msf("com.sun.star.drawing.GraphicObjectShape")
            if gos is None:
                print("Could not create a graphic object shape")
                return
            
            bitmap = Images.get_bitmap(fnm)
            if bitmap is None:
                return
            # store the image's bitmap as the graphic shape's URL's value
            Props.set_property(prop_set=gos, name="GraphicURL", value=bitmap)
            
            # set the shape's size
            xdraw_shape = Lo.qi(XShape, gos)
            xdraw_shape.setSize(im_size)
            
            # insert image shape into the document, followed by newline
            cls._append3(cursor, gos)
            cls.end_line(cursor)
        except Exception as e:
            print(f"Insertion of graphic in '{fnm}' failed:")
            print(f"   {e}")

    @classmethod
    def add_line_divider(cls, cursor: XTextCursor, line_width: int) -> None:
        try:
            ls: XTextContent = Lo.create_instance_msf("com.sun.star.drawing.LineShape")
            if ls is None:
                print("Could not create a line shape")
                return
            
            line_shape = Lo.qi(XShape, ls)
            line_shape.setSize(Size(line_width, 0))
            
            cls.end_paragraph(cursor)
            cls._append3(cursor, ls)
            cls.end_paragraph(cursor)
            
            # center the previous paragraph
            cls.style_prev_paragraph(cursor=cursor, prop_val=PA_CENTER, prop_name="ParaAdjust")
            
            cls.end_paragraph(cursor)
        except Exception as e:
            print("Insertion of graphic line failed")
            print(f"    {e}")
    
    # =================== extracting graphics from text doc ================
    
    @classmethod
    def get_text_graphics(cls, text_doc: XTextDocument) -> List[XGraphic] | None:
        xname_access = cls.get_graphic_links(text_doc)
        if xname_access is None:
            return None
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
                xgraphic = Images.load_graphic_link(graphic_link)
                if xgraphic is not None:
                    pics.append(xgraphic)
                else:
                    print(f"{name} could not be accessed")
        if len(pics) == 0:
            return None
        return pics
    
    @staticmethod
    def get_graphic_links(doc: XComponent) -> XNameAccess | None:
        ims_supplier = Lo.qi(XTextGraphicObjectsSupplier, doc)
        if ims_supplier is None:
            print("Text graphics supplier could not be created")
            return None
        
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
        serv_info = Lo.qi(XServiceInfo, graphic)
        return serv_info is not None and serv_info.supportsService("com.sun.star.text.TextContent") and serv_info.supportsService("com.sun.star.text.TextGraphicObject")

    @staticmethod
    def get_shapes(text_doc: XTextDocument) -> XDrawPage | None:
        draw_page_supplier = Lo.qi(XDrawPageSupplier, text_doc)
        if draw_page_supplier is None:
            print("Draw page supplier could not be created")
            return None
        
        return draw_page_supplier.getDrawPage()
    
    # -----------------  Linguistic API ------------------------

    @classmethod
    def print_services_info(cls, lingo_mgr: XLinguServiceManager2) -> None:
        loc = Locale("en","US","")
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
        service_names = lingo_mgr.getAvailableServices(f"com.sun.star.linguistic2.{service}". loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")
    
    @staticmethod
    def print_config_service_info(lingo_mgr: XLinguServiceManager2, service: str, loc: Locale) -> None:
        service_names = lingo_mgr.getAvailableServices(f"com.sun.star.linguistic2.{service}". loc)
        print(f"{service} ({len(service_names)}):")
        for name in service_names:
            print(f"  {name}")
    
    @staticmethod
    def print_locales(service: str, loc: Iterable[Locale]) -> None:
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
        loc = Locale("en","US","")
        impl_names = (impl_name,)
        lingo_mgr.setConfiguredServices(f"com.sun.star.linguistic2.{service}", loc, impl_names)
    
    @classmethod
    def dicts_info(cls) -> None:
        dict_lst: XSearchableDictionaryList = Lo.create_instance_mcf("com.sun.star.linguistic2.DictionaryList")
        if not dict_lst:
            print("No list of dictionaries found")
            return
        cls.print_dicts_info(dict_lst)
        
        cd_list: XConversionDictionaryList = Lo.create_instance_mcf("com.sun.star.linguistic2.ConversionDictionaryList")
        if cd_list is None:
            print("No list of conversion dictionaries found")
            return
        cls.print_con_dicts_info(cd_list)
    
    @classmethod
    def print_dicts_info(cls, dict_list: XSearchableDictionaryList) -> None:
        if dict_list is None:
            print("Dictionary list is null")
            return
        print(f"No. of dictionaries: {dict_list.getCount()}")
        dicts = dict_list.getDictionaries()
        for d in dicts:
            print(f"  {d.getName()} ({d.getCount()}); ({'active' if d.isActive() else 'na'}); '{d.getLocale().Country}'; {cls.get_dict_type(d.getDictionaryType())}")
        print()
    
    @staticmethod
    def get_dict_type(dt:uno.Enum) -> str:
        if dt == DT_POSITIVE:
            return "positive"
        if dt == DT_NEGATIVE:
            return "negative"
        if dt == DT_MIXED:
            return "mixed"
        return "??"
    
    @staticmethod
    def print_con_dicts_info(cd_lst:XConversionDictionaryList) -> None:
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
        return Lo.create_instance_mcf("com.sun.star.linguistic2.LinguProperties")

    # ---------------- Linguistics: spell checking --------------

    @staticmethod
    def load_spell_checker() -> XSpellChecker | None:
        lingo_mgr: XLinguServiceManager = Lo.create_instance_mcf("com.sun.star.linguistic2.LinguServiceManager")
        if lingo_mgr is None:
            print("No linguistics manager found")
            return None
        return lingo_mgr.getSpellChecker()
    
    @classmethod
    def spell_sentence(cls, sent: str, speller:XSpellChecker) -> int:
        # https://tinyurl.com/y6o8doh2
        words = re.split('\W+', sent)
        count = 0
        for word in words:
            is_correct = cls.spell_word(word, speller)
            count = count + (0 if is_correct else 1)
        return count
    
    @staticmethod
    def spell_word(word: str, speller: XSpellChecker) -> bool:
        loc = Locale("en", "US", "")
        alts = speller.spell(word, loc, tuple())
        if alts is not None:
            print(f"* '{word}' is unknown. Try:")
            alt_words = alts.getAlternatives()
            Lo.print_names(alt_words)
            return False
        return True
    
    # ---------------- Linguistics: thesaurus --------------
    
    @staticmethod
    def load_thesaurus() -> XThesaurus:
        lingo_mgr: XLinguServiceManager = Lo.create_instance_mcf("com.sun.star.linguistic2.LinguServiceManager")
        if lingo_mgr is None:
            print("No linguistics manager found")
            return None
        return lingo_mgr.getThesaurus()
    
    @staticmethod
    def print_meaning(word: str, thesaurus: XThesaurus) -> int:
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
    
    # ---------------- Linguistics: grammar checking --------------
    
    @staticmethod
    def load_proofreader() -> XProofreader:
        return Lo.create_instance_mcf("com.sun.star.linguistic2.Proofreader")
    
    @classmethod
    def proof_sentence(cls, sent: str, proofreader: XProofreader) -> int:
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
    def print_proof_error(string:str, err: SingleProofreadingError) -> None:
        e_end = err.nErrorStart + err.nErrorLength
        err_txt = string[err.nErrorStart:e_end]
        print(f"G* {err.aShortComment} in: '{err_txt}'")
        if err.aSuggestions > 0:
            print(f"  Suggested change: '{err.aSuggestions[0]}'")
        print()
    
    # ---------------- Linguistics: location guessing --------------
    
    @staticmethod
    def guess_locale(test_str: str) -> Locale | None:
        guesser: XLanguageGuessing = Lo.create_instance_mcf("com.sun.star.linguistic2.LanguageGuessing")
        if guesser is None:
            print("No language guesser found")
            return None
        return guesser.guessPrimaryLanguage(test_str, 0, len(test_str))
    
    @staticmethod
    def print_locale(loc: Locale) -> None:
        if loc is not None:
            print(f"Locale lang: '{loc.Language}'; country: '{loc.Country}'; variant: '{loc.Variant}'")
    
    # ---------------- Linguistics dialogs and menu items --------------
    
    @staticmethod
    def open_sent_check_options() -> None:
        """open Options - Language Settings - English sentence checking"""
        pip = Info.get_pip()
        lang_ext = pip.getPackageLocation("org.openoffice.en.hunspell.dictionaries")
        print(f"Lang Ext: {lang_ext}")
        url = f"{lang_ext}/dialog/en.xdl"
        props = Props.make_props(OptionsPageURL=url)
        Lo.dispatch_cmd(cmd="OptionsTreeDialog", props=props)
        Lo.wait(2000)
    
    @staticmethod
    def open_spell_grammar_dialog() -> None:
        """activate dialog in  Tools > Speling and Grammar..."""
        Lo.dispatch_cmd("SpellingAndGrammarDialog")
        Lo.wait(2000)
    
    @staticmethod
    def toggle_auto_spell_check() -> None:
        Lo.dispatch_cmd("SpellOnline")
    
    @staticmethod
    def open_thesaurus_dialog() -> None:
        Lo.dispatch_cmd("ThesaurusDialog")
    