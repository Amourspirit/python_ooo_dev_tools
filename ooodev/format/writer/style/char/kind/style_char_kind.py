from enum import Enum


class StyleCharKind(Enum):
    """Style Look ups for Character Styles"""

    BULLET_SYMBOLS = "Bullet Symbols"
    CAPTION_CHARACTERS = "Caption characters"
    CITATION = "Citation"
    DEFINITION = "Definition"
    DROP_CAPS = "Drop Caps"
    EMPHASIS = "Emphasis"
    ENDNOTE_SYMBOL = "Endnote Symbol"
    ENDNOTE_ANCHOR = "Endnote anchor"
    EXAMPLE = "Example"
    FOOTNOTE_SYMBOL = "Footnote Symbol"
    FOOTNOTE_ANCHOR = "Footnote anchor"
    INDEX_LINK = "Index Link"
    INTERNET_LINK = "Internet link"
    LINE_NUMBERING = "Line numbering"
    MAIN_INDEX_ENTRY = "Main index entry"
    NUMBERING_SYMBOLS = "Numbering Symbols"
    PAGE_NUMBER = "Page Number"
    PLACEHOLDER = "Placeholder"
    RUBIES = "Rubies"
    SOURCE_TEXT = "Source Text"
    STANDARD = "Standard"
    """Removes Character Styling"""
    STRONG_EMPHASIS = "Strong Emphasis"
    TELETYPE = "Teletype"
    USER_ENTRY = "User Entry"
    VARIABLE = "Variable"
    VERTICAL_NUMBERING_SYMBOLS = "Vertical Numbering Symbols"
    VISITED_INTERNET_LINK = "Visited Internet Link"

    def __str__(self) -> str:
        return self.value
