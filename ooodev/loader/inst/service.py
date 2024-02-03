from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from enum import Enum

if TYPE_CHECKING:
    from ooodev.loader.inst.doc_type import DocType, DocTypeStr


class Service(str, Enum):
    """Service Type"""

    UNKNOWN = "com.sun.frame.XModel"
    WRITER = "com.sun.star.text.TextDocument"
    BASE = "com.sun.star.sdb.OfficeDatabaseDocument"
    CALC = "com.sun.star.sheet.SpreadsheetDocument"
    DRAW = "com.sun.star.drawing.DrawingDocument"
    IMPRESS = "com.sun.star.presentation.PresentationDocument"
    MATH = "com.sun.star.formula.FormulaProperties"

    def __str__(self) -> str:
        return self.value

    def get_doc_type(self) -> DocType:
        """Gets the document type as DocType Enum"""
        from ooodev.loader.inst.doc_type import DocType

        return DocType[self.name]

    def get_doc_type_str(self) -> DocTypeStr:
        """Gets the document type as DocTypeStr Enum"""
        from ooodev.loader.inst.doc_type import DocTypeStr

        return DocTypeStr[self.name]
