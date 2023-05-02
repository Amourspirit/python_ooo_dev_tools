from __future__ import annotations
import uno
from enum import Enum


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
