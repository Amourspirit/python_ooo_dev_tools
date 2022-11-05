from enum import Enum
from . import kind_helper


class DrawingHatchingKind(str, Enum):
    """
    Drawing (some) Hatching Values.
    The actual value may change from version to version and may be different
    in other languages.

    These are just quick suggestion that show up in Draw Hatching options.
    """

    BLACK_0_DEGREES = "Black 0 Degrees"
    BLACK_180_DEGREES_CROSSED = "Black 180 Degrees Crossed"
    BLACK_90_DEGREES = "Black 90 Degrees"
    BLUE_45_DEGREES = "Blue 45 Degrees"
    BLUE_45_DEGREES_CROSSED = "Blue 45 Degrees Crossed"
    BLUE_NEG_45_DEGREES = "Blue -45 Degrees"
    GREEN_30_DEGREES = "Green 30 Degrees"
    GREEN_60_DEGREES = "Green 60 Degrees"
    GREEN_90_DEGREES_TRIPLE = "Green 90 Degrees Triple"
    RED_45_DEGREES = "Red 45 Degrees"
    RED_90_DEGREES_CROSSED = "Red 90 Degrees Crossed"
    RED_NEG_45_DEGREES_TRIPLE = "Red -45 Degrees Triple"
    YELLOW_45_DEGREES = "Yellow 45 Degrees"
    YELLOW_45_DEGREES_CROSSED = "Yellow 45 Degrees Crossed"
    YELLOW_45_DEGREES_TRIPLE = "Yellow 45 Degrees Triple"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "DrawingHatchingKind":
        """
        Gets an ``DrawingHatchingKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingHatchingKind`` instance.

        Returns:
            DrawingHatchingKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingHatchingKind)
