from enum import IntEnum
from ooodev.utils.kind import kind_helper


class DateFormatKind(IntEnum):
    """Specifies the format of the displayed date."""

    SYSTEM_SHORT = 0
    """Standard (short)"""
    SYSTEM_SHORT_YY = 1
    """Standard (short YY)"""
    SYSTEM_SHORT_YYYY = 2
    """Standard (short YYYY)"""
    SYSTEM_LONG = 3
    """Standard (long)"""
    DD_MM_YY = 4
    """DD/MM/YY"""
    MM_DD_YY = 5
    """MM/DD/YY"""
    YY_MM_DD = 6
    """YY/MM/DD"""
    DD_MM_YYYY = 7
    """DD/MM/YYYY"""
    MM_DD_YYYY = 8
    """MM/DD/YYYY"""
    YYYY_MM_DD = 9
    """YYYY/MM/DD"""
    DIN_5008_YY_MM_DD = 10
    """YY-MM-DD"""
    DIN_5008_YYYY_MM_DD = 11
    """YYYY-MM-DD"""

    @staticmethod
    def from_str(s: str) -> "DateFormatKind":
        """
        Gets an ``DateFormatKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DateFormatKind`` instance.

        Returns:
            DateFormatKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DateFormatKind)
