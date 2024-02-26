from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaHtmlKind(Enum):
    """Style Look ups for Paragraph HTML Styles"""

    ENDNOTE = StyleParaKind.ENDNOTE.value
    FOOTNOTE = StyleParaKind.FOOTNOTE.value
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
    HORIZONTAL_LINE = StyleParaKind.HORIZONTAL_LINE.value
    LIST_CONTENTS = StyleParaKind.LIST_CONTENTS.value
    LIST_HEADING = StyleParaKind.LIST_HEADING.value
    PREFORMATTED_TEXT = StyleParaKind.PREFORMATTED_TEXT.value
    QUOTATIONS = StyleParaKind.QUOTATIONS.value
    SENDER = StyleParaKind.SENDER.value
    TABLE_CONTENTS = StyleParaKind.TABLE_CONTENTS.value
    TABLE_HEADING = StyleParaKind.TABLE_HEADING.value
    TEXT_BODY = StyleParaKind.TEXT_BODY.value

    def __str__(self) -> str:
        return self.value
