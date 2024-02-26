from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaIndexKind(Enum):
    """Style Look ups for Paragraph Index Styles"""

    BIBLIOGRAPHY_1 = StyleParaKind.BIBLIOGRAPHY_1.value
    BIBLIOGRAPHY_HEADING = StyleParaKind.BIBLIOGRAPHY_HEADING.value
    CONTENTS_1 = StyleParaKind.CONTENTS_1.value
    CONTENTS_10 = StyleParaKind.CONTENTS_10.value
    CONTENTS_2 = StyleParaKind.CONTENTS_2.value
    CONTENTS_3 = StyleParaKind.CONTENTS_3.value
    CONTENTS_4 = StyleParaKind.CONTENTS_4.value
    CONTENTS_5 = StyleParaKind.CONTENTS_5.value
    CONTENTS_6 = StyleParaKind.CONTENTS_6.value
    CONTENTS_7 = StyleParaKind.CONTENTS_7.value
    CONTENTS_8 = StyleParaKind.CONTENTS_8.value
    CONTENTS_9 = StyleParaKind.CONTENTS_9.value
    CONTENTS_HEADING = StyleParaKind.CONTENTS_HEADING.value
    FIGURE_INDEX_1 = StyleParaKind.FIGURE_INDEX_1.value
    FIGURE_INDEX_HEADING = StyleParaKind.FIGURE_INDEX_HEADING.value
    INDEX = StyleParaKind.INDEX.value
    INDEX_1 = StyleParaKind.INDEX_1.value
    INDEX_2 = StyleParaKind.INDEX_2.value
    INDEX_3 = StyleParaKind.INDEX_3.value
    INDEX_HEADING = StyleParaKind.INDEX_HEADING.value
    INDEX_SEPARATOR = StyleParaKind.INDEX_SEPARATOR.value
    OBJECT_INDEX_1 = StyleParaKind.OBJECT_INDEX_1.value
    OBJECT_INDEX_HEADING = StyleParaKind.OBJECT_INDEX_HEADING.value
    TABLE_INDEX_1 = StyleParaKind.TABLE_INDEX_1.value
    TABLE_INDEX_HEADING = StyleParaKind.TABLE_INDEX_HEADING.value
    USER_INDEX_1 = StyleParaKind.USER_INDEX_1.value
    USER_INDEX_10 = StyleParaKind.USER_INDEX_10.value
    USER_INDEX_2 = StyleParaKind.USER_INDEX_2.value
    USER_INDEX_3 = StyleParaKind.USER_INDEX_3.value
    USER_INDEX_4 = StyleParaKind.USER_INDEX_4.value
    USER_INDEX_5 = StyleParaKind.USER_INDEX_5.value
    USER_INDEX_6 = StyleParaKind.USER_INDEX_6.value
    USER_INDEX_7 = StyleParaKind.USER_INDEX_7.value
    USER_INDEX_8 = StyleParaKind.USER_INDEX_8.value
    USER_INDEX_9 = StyleParaKind.USER_INDEX_9.value
    USER_INDEX_HEADING = StyleParaKind.USER_INDEX_HEADING.value

    def __str__(self) -> str:
        return self.value
