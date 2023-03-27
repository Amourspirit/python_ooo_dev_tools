from enum import Enum


class WriterStylePageKind(Enum):
    """Style Look ups for Page Styles"""

    ENDNOTE = "Endnote"
    ENVELOPE = "Envelope"
    FIRST_PAGE = "First Page"
    FOOTNOTE = "Footnote"
    HTML = "HTML"
    INDEX = "Index"
    LANDSCAPE = "Landscape"
    LEFT_PAGE = "Left Page"
    RIGHT_PAGE = "Right Page"
    STANDARD = "Standard"

    def __str__(self) -> str:
        return self.value
