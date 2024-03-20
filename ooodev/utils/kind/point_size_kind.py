from enum import IntEnum
from ooodev.utils.kind import kind_helper


class PointSizeKind(IntEnum):
    """
    PointSizeKind Enum class.

    .. versionadded:: 0.36.3
    """

    X = 0
    """Represents X of a Point"""
    Y = 1
    """Represents Y of a Point"""
    WIDTH = 2
    """Represents Width of a Size"""
    HEIGHT = 3
    """Represents Height of a Size"""

    @staticmethod
    def from_str(s: str) -> "PointSizeKind":
        """
        Gets an ``PointSizeKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``PointSizeKind`` instance.

        Returns:
            PointSizeKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, PointSizeKind)
