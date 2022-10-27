from enum import Enum


class LineStyleNameKind(str, Enum):
    """
    Line Style Name Kind.

    Used to encapsulate Line Style Names.
    """

    DASHED = "Dashed"
    DASHES_3_DOTS_3 = "3 Dashes 3 Dots"
    DOTS_2_DASH_1 = "2 Dots 1 Dash"
    DOTS_2_DASHES_3 = "2 Dots 3 Dashes"
    FINE_DASHED = "Fine Dashed"
    FINE_DOTTED = "Fine Dotted"
    LINE_STYLE_9 = "Line Style 9"
    LINE_WITH_FINE_DOTS = "Line with Fine Dots"
    ULTRA_FINE_DOTTED = "Ultrafine Dotted"
    ULTRAFINE_DASHED = "Ultrafine Dashed"

    def __str__(self) -> str:
        return self.value
