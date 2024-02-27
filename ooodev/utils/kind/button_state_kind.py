from enum import IntEnum
from ooodev.utils.kind import kind_helper


class ButtonStateKind(IntEnum):
    """
    Specifies state of a tri-state button control.

    .. versionadded:: 0.29.0
    """

    NOT_PRESSED = 0
    """State not pressed"""
    PRESSED = 1
    """State pressed"""
    DONT_KNOW = 2
    """State don't know"""

    @staticmethod
    def from_str(s: str) -> "ButtonStateKind":
        """
        Gets an ``ButtonStateKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``ButtonStateKind`` instance.

        Returns:
            ButtonStateKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ButtonStateKind)
