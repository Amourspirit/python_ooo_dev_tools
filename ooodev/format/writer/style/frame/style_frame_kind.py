from enum import Enum


class StyleFrameKind(Enum):
    """Style Look ups for Frame Styles"""

    FORMULA = "Formula"
    FRAME = "Frame"
    GRAPHICS = "Graphics"
    LABELS = "Labels"
    MARGINALIA = "Marginalia"
    OLE = "OLE"
    WATERMARK = "Watermark"

    def __str__(self) -> str:
        return self.value
