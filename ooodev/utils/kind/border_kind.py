from enum import IntEnum
from ooodev.utils.kind import kind_helper


class BorderKind(IntEnum):
    """Specifies the border style of the control."""

    NONE = 0
    """No border"""
    BORDER_3D = 1
    """3D border"""
    BORDER_SIMPLE = 2
    """Simple border"""

    @staticmethod
    def from_str(s: str) -> "BorderKind":
        """
        Gets an ``BorderKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``BorderKind`` instance.

        Returns:
            BorderKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, BorderKind)
