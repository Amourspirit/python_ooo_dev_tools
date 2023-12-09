from enum import Enum


class TextPortionTypeKind(Enum):
    """The type of a text portion."""

    # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextPortion.html#a7ecd2de53df4ec8d3fffa94c2e80d651
    TEXT = "Text"
    """string content"""
    TEXT_FIELD = "TextField"
    """a text field"""
    TEXT_CONTENT = "TextContent"
    """Text content - supplied via the interface com.sun.star.container.XContentEnumerationAccess"""
    CONTROL_CHARACTER = "ControlCharacter"
    """A control character"""
    FOOTNOTE = "Footnote"
    """A footnote or an endnote"""
    REFERENCE_MARK = "ReferenceMark"
    """a reference mark"""
    DOCUMENT_INDEX_MARK = "DocumentIndexMark"
    """A document index mark"""
    BOOKMARK = "Bookmark"
    """A bookmark"""
    REDLINE = "Redline"
    """A redline portion which is a result of the change tracking feature"""
    RUBY = "Ruby"
    """A ruby attribute which is used in Asian text"""
    FRAME = "Frame"
    """A frame"""
    SOFT_PAGE_BREAK = "SoftPageBreak"
    """A soft page break"""
    IN_CONTENT_METADATA = "InContentMetadata"
    """A text range with attached metadata"""
    UNKNOWN = "unknown"

    def __str__(self) -> str:
        return self.value
