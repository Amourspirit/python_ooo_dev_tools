from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaTextKind(Enum):
    """Style Look ups for Paragraph Text Styles"""

    FIRST_LINE_INDENT = StyleParaKind.FIRST_LINE_INDENT.value
    HANGING_INDENT = StyleParaKind.HANGING_INDENT.value
    HEADER = StyleParaKind.HEADER.value
    HEADER_AND_FOOTER = StyleParaKind.HEADER_FOOTER.value
    HEADER_LEFT = StyleParaKind.HEADER_LEFT.value
    HEADER_RIGHT = StyleParaKind.HEADER_RIGHT.value
    HEADING = StyleParaKind.HEADING.value
    HEADING_1 = StyleParaKind.HEADING_1.value
    HEADING_10 = StyleParaKind.HEADING_10.value
    HEADING_2 = StyleParaKind.HEADING_2.value
    HEADING_3 = StyleParaKind.HEADING_3.value
    HEADING_4 = StyleParaKind.HEADING_4.value
    HEADING_5 = StyleParaKind.HEADING_5.value
    HEADING_6 = StyleParaKind.HEADING_6.value
    HEADING_7 = StyleParaKind.HEADING_7.value
    HEADING_8 = StyleParaKind.HEADING_8.value
    HEADING_9 = StyleParaKind.HEADING_9.value
    LIST_INDENT = StyleParaKind.LIST_INDENT.value
    MARGINALIA = StyleParaKind.SIGNATURE.value
    SIGNATURE = StyleParaKind.TEXT_BODY.value
    TEXT_BODY = StyleParaKind.TEXT_BODY.value
    TEXT_BODY_INDENT = StyleParaKind.TEXT_BODY_INDENT.value

    def __str__(self) -> str:
        return self.value
