from enum import IntEnum
from ooodev.utils.kind import kind_helper


class TimeFormatKind(IntEnum):
    """Specifies the format of the displayed time."""

    SHORT_24H = 0
    """24h short"""
    LONG_24H = 1
    """12h short"""
    SHORT_12H = 2
    """12h short"""
    LONG_12H = 3
    """12h long"""
    DURATION_SHORT = 4
    """Duration short"""
    DURATION_LONG = 5
    """Duration long"""

    @staticmethod
    def from_str(s: str) -> "TimeFormatKind":
        """
        Gets an ``TimeFormatKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``TimeFormatKind`` instance.

        Returns:
            TimeFormatKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, TimeFormatKind)
