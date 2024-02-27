from enum import IntEnum
from ooodev.utils.kind import kind_helper


class TriStateKind(IntEnum):
    """
    Specifies state of a tri-state control such as a check box.
    """

    NOT_CHECKED = 0
    """State not checked"""
    CHECKED = 1
    """State checked"""
    DONT_KNOW = 2
    """State don't know"""

    @staticmethod
    def from_str(s: str) -> "TriStateKind":
        """
        Gets an ``TriStateKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``TriStateKind`` instance.

        Returns:
            TriStateKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, TriStateKind)
