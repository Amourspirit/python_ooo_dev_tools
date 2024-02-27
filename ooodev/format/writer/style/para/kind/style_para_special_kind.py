from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaSpecialKind(Enum):
    """Style Look ups for Paragraph Special Styles"""

    ADDRESSEE = StyleParaKind.ADDRESSEE.value
    CAPTION = StyleParaKind.CAPTION.value
    DRAWING = StyleParaKind.DRAWING.value
    ENDNOTE = StyleParaKind.ENDNOTE.value
    FIGURE = StyleParaKind.FIGURE.value
    FOOTER = StyleParaKind.FOOTER.value
    FOOTER_LEFT = StyleParaKind.FOOTER_LEFT.value
    FOOTER_RIGHT = StyleParaKind.FOOTER_RIGHT.value
    FOOTNOTE = StyleParaKind.FOOTNOTE.value
    FRAME_CONTENTS = StyleParaKind.FRAME_CONTENTS.value
    HEADER = StyleParaKind.HEADER.value
    HEADER_FOOTER = StyleParaKind.HEADER_FOOTER.value
    HEADER_LEFT = StyleParaKind.HEADER_LEFT.value
    HEADER_RIGHT = StyleParaKind.HEADER_RIGHT.value
    ILLUSTRATION = StyleParaKind.ILLUSTRATION.value
    SENDER = StyleParaKind.SENDER.value
    TABLE = StyleParaKind.TABLE.value
    TABLE_CONTENTS = StyleParaKind.TABLE_CONTENTS.value
    TABLE_HEADING = StyleParaKind.TABLE_HEADING.value
    TEXT = StyleParaKind.TEXT.value

    def __str__(self) -> str:
        return self.value
