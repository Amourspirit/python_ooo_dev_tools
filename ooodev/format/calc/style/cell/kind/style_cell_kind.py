from enum import Enum


class StyleCellKind(Enum):
    """Style Lookup for Cell Styles"""

    ACCENT = "Accent"
    ACCENT_1 = "Accent 1"
    ACCENT_2 = "Accent 2"
    ACCENT_3 = "Accent 3"
    BAD = "Bad"
    DEFAULT = "Default"
    ERROR = "Error"
    FOOTNOTE = "Footnote"
    GOOD = "Good"
    HEADING = "Heading"
    HEADING_1 = "Heading 1"
    HEADING_2 = "Heading 2"
    HYPERLINK = "Hyperlink"
    NEUTRAL = "Neutral"
    NOTE = "Note"
    RESULT = "Result"
    STATUS = "Status"
    TEXT = "Text"
    WARNING = "Warning"

    def __str__(self) -> str:
        return self.value
