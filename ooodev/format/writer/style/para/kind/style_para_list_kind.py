from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaListKind(Enum):
    """Style Look ups for Paragraph List Styles"""

    LIST = StyleParaKind.LIST.value
    LIST_1 = StyleParaKind.LIST_1.value
    LIST_1_CONT = StyleParaKind.LIST_1_CONT.value
    LIST_1_END = StyleParaKind.LIST_1_END.value
    LIST_1_START = StyleParaKind.LIST_1_START.value
    LIST_2 = StyleParaKind.LIST_2.value
    LIST_2_CONT = StyleParaKind.LIST_2_CONT.value
    LIST_2_END = StyleParaKind.LIST_2_END.value
    LIST_2_START = StyleParaKind.LIST_2_START.value
    LIST_3 = StyleParaKind.LIST_3.value
    LIST_3_CONT = StyleParaKind.LIST_3_CONT.value
    LIST_3_END = StyleParaKind.LIST_3_END.value
    LIST_3_START = StyleParaKind.LIST_3_START.value
    LIST_4 = StyleParaKind.LIST_4.value
    LIST_4_CONT = StyleParaKind.LIST_4_CONT.value
    LIST_4_END = StyleParaKind.LIST_4_END.value
    LIST_4_START = StyleParaKind.LIST_4_START.value
    LIST_5 = StyleParaKind.LIST_5.value
    LIST_5_CONT = StyleParaKind.LIST_5_CONT.value
    LIST_5_END = StyleParaKind.LIST_5_END.value
    LIST_5_START = StyleParaKind.LIST_5_START.value
    LIST_CONTENTS = StyleParaKind.LIST_CONTENTS.value
    LIST_HEADING = StyleParaKind.LIST_HEADING.value
    LIST_INDENT = StyleParaKind.LIST_INDENT.value
    NUMBERING_1 = StyleParaKind.NUMBERING_1.value
    NUMBERING_1_CONT = StyleParaKind.NUMBERING_1_CONT.value
    NUMBERING_1_END = StyleParaKind.NUMBERING_1_END.value
    NUMBERING_1_START = StyleParaKind.NUMBERING_1_START.value
    NUMBERING_2 = StyleParaKind.LIST.value
    NUMBERING_2_CONT = StyleParaKind.NUMBERING_2_CONT.value
    NUMBERING_2_END = StyleParaKind.NUMBERING_2_END.value
    NUMBERING_2_START = StyleParaKind.NUMBERING_2_START.value
    NUMBERING_3 = StyleParaKind.NUMBERING_3.value
    NUMBERING_3_CONT = StyleParaKind.NUMBERING_3_CONT.value
    NUMBERING_3_END = StyleParaKind.NUMBERING_3_END.value
    NUMBERING_3_START = StyleParaKind.NUMBERING_3_START.value
    NUMBERING_4 = StyleParaKind.NUMBERING_4.value
    NUMBERING_4_CONT = StyleParaKind.NUMBERING_4_CONT.value
    NUMBERING_4_END = StyleParaKind.NUMBERING_4_END.value
    NUMBERING_4_START = StyleParaKind.NUMBERING_4_START.value
    NUMBERING_5 = StyleParaKind.NUMBERING_5.value
    NUMBERING_5_CONT = StyleParaKind.NUMBERING_5_CONT.value
    NUMBERING_5_END = StyleParaKind.NUMBERING_5_END.value
    NUMBERING_5_START = StyleParaKind.NUMBERING_5_START.value

    def __str__(self) -> str:
        return self.value
