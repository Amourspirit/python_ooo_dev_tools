# coding: utf-8
from __future__ import annotations

from typing import Union, TYPE_CHECKING, overload, cast

from ..utils import lo as m_lo
from ..utils import info as m_info
from ..utils import file_io as m_file_io
from ..utils import props as m_props

from com.sun.star.awt import FontWeight # type: ignore
from com.sun.star.awt.FontSlant import ITALIC as FS_ITALIC # type: ignore
from com.sun.star.awt import Size
from com.sun.star.container import XEnumerationAccess
from com.sun.star.document import XDocumentInsertable
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.style import NumberingType # type: ignore
from com.sun.star.style.BreakType import PAGE_AFTER as BT_PAGE_AFTER, COLUMN_AFTER as BT_COLUMN_AFTER # type: ignore
from com.sun.star.style.ParagraphAdjust import CENTER as PA_CENTER, RIGHT as PA_RIGHT # type: ignore
from com.sun.star.text import ControlCharacter # type: ignore
from com.sun.star.text import XPageCursor
from com.sun.star.text import XParagraphCursor
from com.sun.star.text import XSentenceCursor
from com.sun.star.text import XTextContent
from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextViewCursor
from com.sun.star.text import XWordCursor
from com.sun.star.text.PageNumberType import CURRENT as PN_CURRENT # type: ignore
from com.sun.star.view.PaperFormat import A4 as PF_A4 # type: ignore

if TYPE_CHECKING:
    from com.sun.star.container import XEnumeration
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XModel
    from com.sun.star.lang import XComponent
    from com.sun.star.beans import XPropertySet
    from com.sun.star.text import XText
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextField
    from com.sun.star.text import XTextViewCursorSupplier
    from com.sun.star.view import XPrintable

Lo = m_lo.Lo
Info = m_info.Info
FileIO = m_file_io.FileIO
Props = m_props.Props


class Write:
    
    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader) -> Union[XTextDocument, None]:
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
    def open_flat_doc_using_text_template(cls, fnm: str, template_path: str, loader: XComponentLoader) -> Union[XTextDocument, None]:
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
        cursor: Union[XTextCursor, XDocumentInsertable] = cls.get_cursor(doc)
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
    def get_cursor(cursor_obj: XTextDocument) -> Union[XTextCursor, None]:...

    @overload
    @staticmethod
    def get_cursor(cursor_obj: XTextViewCursor) -> Union[XTextCursor, None]:...
    
    
    @staticmethod
    def get_cursor(cursor_obj: Union[XTextDocument, XTextViewCursor]) -> Union[XTextCursor, None]:
        xtext = cursor_obj.getText()
        if xtext is None:
            print("Text not found in document")
            return None
        if Info.support_service(cursor_obj, XTextViewCursor):
            return xtext.createTextCursorByRange(cursor_obj)
        return xtext.createTextCursor()
    
    @classmethod
    def get_word_cursor(cls, text_doc: XTextDocument) -> Union[XWordCursor, None]:
        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            print("Text cursor is null")
            return None
        if not Info.support_service(cursor, XWordCursor):
            print("Text is not XWordCursor")
            return None
        return cursor
    
    @classmethod
    def get_sentence_cursor(cls, text_doc: XTextDocument) -> Union[XSentenceCursor, None]:
        cursor = cls.get_cursor(text_doc)
        if cursor is None:
            print("Text cursor is null")
            return None
        if not Info.support_service(cursor, XSentenceCursor):
            print("Text is not XSentenceCursor")
            return None
        return cursor
    
    @classmethod
    def get_paragraph_cursor(cls, text_doc: XTextDocument) -> Union[XParagraphCursor, None]:
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
    def get_view_cursor(text_doc: Union[XTextDocument, XModel]) -> XTextViewCursor:
        xcontroller: XTextViewCursorSupplier = text_doc.getCurrentController()
        return xcontroller.getViewCursor()
    
    @staticmethod
    def get_current_page(tv_cursor: Union[XTextViewCursor, XPageCursor]) -> int:
        if not Info.support_service(tv_cursor, XPageCursor):
            print("Could not create a page cursor")
            return -1
        return tv_cursor.getPage()
    
    @staticmethod
    def get_coord_str(tv_cursor: XTextViewCursor) -> str:
        pos = tv_cursor.getPosition()
        return f"({pos.X}, {pos.Y})"
    
    @staticmethod
    def get_page_count(text_doc: Union[XTextDocument, XModel]) -> int:
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
    def get_enumeration(obj: XEnumerationAccess) -> Union[XEnumeration, None]:
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
    def style_prev_paragraph(cursor: Union[XTextCursor, XParagraphCursor], prop_val: object, prop_name: str = None) -> None:
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
    def get_page_size(text_doc: XTextDocument) -> Union[Size, None]:
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
    def set_a4_page_format(text_doc: Union[XTextDocument, XPrintable]) -> None:
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
            cls.append(cursor=footer_cursor,text_content=cls.get_page_number())
            cls.append(cursor=footer_cursor, text=" of ")
            cls.append(cursor=footer_cursor,text_content=cls.get_page_count())
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
            header_cursor = cast(Union[XTextCursor, XPropertySet], header_text.createTextCursor())
            header_cursor.gotoEnd(False)
            
            header_cursor.setPropertyValue("CharFontName", Info.get_font_general_name())
            header_cursor.setPropertyValue("CharHeight", 10)
            header_cursor.setPropertyValue("ParaAdjust", PA_RIGHT)
            
            header_text.setString(f"{h_text}\n")
        except Exception as e:
            print(e)

    @staticmethod
    def get_draw_page(doc: Union[XTextDocument, XDrawPageSupplier]) -> Union[XDrawPage, None]:
        if not Info.support_service(obj=doc, service=XDrawPageSupplier):
            print("XDrawPageSupplier interface is not supporte for doc")
            return None
        return doc.getDrawPage()

    # -------------------------- adding elements ----------------------------
    