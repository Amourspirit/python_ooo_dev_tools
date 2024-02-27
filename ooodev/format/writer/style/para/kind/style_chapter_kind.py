from enum import Enum
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


class StyleChapterKind(Enum):
    """Style Look ups for Paragraph Chapter Styles"""

    APPENDIX = StyleParaKind.APPENDIX.value
    SUBTITLE = StyleParaKind.SUBTITLE.value
    TITLE = StyleParaKind.TITLE.value

    def __str__(self) -> str:
        return self.value
