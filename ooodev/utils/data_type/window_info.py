from __future__ import annotations
import uno
from com.sun.star.lang import XComponent
from com.sun.star.frame import XFrame

from ooodev.loader.inst.doc_type import DocType


class WindowInfo:
    __slots__ = (
        "component",
        "component",
        "frame",
        "window_name",
        "window_title",
        "window_file_name",
        "document_type",
    )

    def __init__(self) -> None:
        self.component: XComponent | None = None  # com.sun.star.lang.XComponent
        self.frame: XFrame | None = None  # com.sun.star.comp.framework.Frame
        self.window_name: str = ""  # Object Name
        self.window_title: str = ""  # Only mean to identify new documents
        self.window_file_name: str = ""  # URL of file name
        self.document_type: DocType = DocType.UNKNOWN  # Writer, Calc, ...
