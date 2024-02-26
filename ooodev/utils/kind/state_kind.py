from enum import IntEnum
from ooodev.utils.kind import kind_helper


class StateKind(IntEnum):
    """
    Specifies state of a state control such as a radio button.
    """

    NOT_CHECKED = 0
    """State not checked"""
    CHECKED = 1
    """State checked"""

    @staticmethod
    def from_str(s: str) -> "StateKind":
        """
        Gets an ``StateKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``StateKind`` instance.

        Returns:
            StateKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, StateKind)
