from enum import Enum
from .style_para_kind import StyleParaKind


class StyleParaChKind(Enum):
    """Style Lookups for Paragraph Chapter Styles"""

    APPENDIX = StyleParaKind.APPENDIX.value
    SUBTITLE = StyleParaKind.SUBTITLE.value
    TITLE = StyleParaKind.TITLE.value

    def __str__(self) -> str:
        return self.value
