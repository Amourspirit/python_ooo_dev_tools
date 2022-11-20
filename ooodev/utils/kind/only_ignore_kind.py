from enum import IntEnum
from . import kind_helper


class OnlyIgnore(IntEnum):
    NONE = 0
    ONLY = 1
    IGNORE = 2

    @staticmethod
    def from_str(s: str) -> "OnlyIgnore":
        """
        Gets an ``OnlyIgnore`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name. Case insensitive.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``OnlyIgnore`` instance.

        Returns:
            OnlyIgnore: Enum instance.
        """
        return kind_helper.enum_from_string(s, OnlyIgnore)
