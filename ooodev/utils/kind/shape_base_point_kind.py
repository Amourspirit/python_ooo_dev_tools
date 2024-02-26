from enum import IntEnum
from ooodev.utils.kind import kind_helper


class ShapeBasePointKind(IntEnum):
    """
    Represent the 9 base points of a shape

    .. versionadded:: 0.9.4
    """

    TOP_LEFT = 1
    """Top Left"""
    TOP_CENTER = 2
    """Top Center"""
    TOP_RIGHT = 3
    """Top Right"""
    CENTER_LEFT = 4
    """Center Left"""
    CENTER = 5
    """Center"""
    CENTER_RIGHT = 6
    """Center Right"""
    BOTTOM_LEFT = 7
    """Bottom Left"""
    BOTTOM_CENTER = 8
    """Bottom Center"""
    BOTTOM_RIGHT = 9
    """Bottom Right"""

    @staticmethod
    def from_str(s: str) -> "ShapeBasePointKind":
        """
        Gets an ``ShapeBasePointKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``ShapeBasePointKind`` instance.

        Returns:
            ShapeBasePointKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ShapeBasePointKind)
