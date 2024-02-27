from enum import Enum
from ooodev.utils.kind import kind_helper


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

    @staticmethod
    def from_str(s: str) -> "LineStyleNameKind":
        """
        Gets an ``LineStyleNameKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``LineStyleNameKind`` instance.

        Returns:
            LineStyleNameKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, LineStyleNameKind)
