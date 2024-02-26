from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleParaCondKind(Enum):
    """Style Look ups for Paragraph Conditional Styles"""

    TEXT_BODY = StyleParaKind.TEXT_BODY.value

    def __str__(self) -> str:
        return self.value
