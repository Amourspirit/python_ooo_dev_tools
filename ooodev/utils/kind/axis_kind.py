from enum import IntEnum
from ooodev.utils.kind import kind_helper


class AxisKind(IntEnum):
    X = 0
    """Represents X Axis"""
    Y = 1
    """Represents Y Axis"""
    Z = 2
    """Represents Z Axis"""

    @staticmethod
    def from_str(s: str) -> "AxisKind":
        """
        Gets an ``AxisKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``AxisKind`` instance.

        Returns:
            AxisKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, AxisKind)
