from enum import IntEnum
from ooodev.utils.kind import kind_helper


class OrientationKind(IntEnum):
    """Specifies the orientation of the control."""

    HORIZONTAL = 0
    """Horizontal orientation"""
    VERTICAL = 1
    """Vertical orientation"""

    @staticmethod
    def from_str(s: str) -> "OrientationKind":
        """
        Gets an ``OrientationKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``OrientationKind`` instance.

        Returns:
            OrientationKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, OrientationKind)
