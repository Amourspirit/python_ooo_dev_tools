from enum import IntEnum
from ooodev.utils.kind import kind_helper


class AlignKind(IntEnum):
    """Specifies the horizontal alignment of the text in the control."""

    LEFT = 0
    """Align left"""
    CENTER = 1
    """Align center"""
    RIGHT = 2
    """Align right"""

    @staticmethod
    def from_str(s: str) -> "AlignKind":
        """
        Gets an ``AlignKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``AlignKind`` instance.

        Returns:
            AlignKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, AlignKind)
