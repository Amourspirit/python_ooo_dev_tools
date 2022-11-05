from enum import IntEnum
from . import kind_helper


class GluePointsKind(IntEnum):
    """Glue Point Kind"""

    TOP = 0
    RIGHT = 1
    BOTTOM = 2
    LEFT = 3

    @staticmethod
    def from_str(s: str) -> "GluePointsKind":
        """
        Gets an ``GluePointsKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``GluePointsKind`` instance.

        Returns:
            GluePointsKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, GluePointsKind)
