from enum import Enum
from ooodev.utils.kind import kind_helper


class DrawingGradientKind(str, Enum):
    """
    Drawing (some) Gradient Values.
    The actual value may change from version to version and may be different
    in other languages.

    These are just quick suggestion that show up in Draw Gradient colors.
    """

    BLANK_WITH_GRAY = "Blank with Gray"
    BLUE_TOUCH = "Blue Touch"
    DEEP_OCEAN = "Deep Ocean"
    GREEN_GRASS = "Green Grass"
    LONDON_MIST = "London Mist"
    MAHOGANY = "Mahogany"
    MIDNIGHT = "Midnight"
    NEON_LIGHT = "Neon Light"
    PASTEL_BOUQUET = "Pastel Bouquet"
    PASTEL_DREAM = "Pastel Dream"
    PRESENT = "Present"
    SPOTTED_GRAY = "Spotted Gray"
    SUBMARINE = "Submarine"
    SUNSHINE = "Sunshine"
    TEAL_TO_BLUE = "Teal to Blue"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "DrawingGradientKind":
        """
        Gets an ``DrawingGradientKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingGradientKind`` instance.

        Returns:
            DrawingGradientKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingGradientKind)
