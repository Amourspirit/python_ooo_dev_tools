from enum import Enum


class StyleFrameKind(Enum):
    """Style Lookups for Frame Styles"""

    FORMULA: str = "Formula"
    FRAME: str = "Frame"
    GRAPHICS: str = "Graphics"
    LABELS: str = "Labels"
    MARGINALIA: str = "Marginalia"
    OLE: str = "OLE"
    WATERMARK: str = "Watermark"

    def __str__(self) -> str:
        return self.value
