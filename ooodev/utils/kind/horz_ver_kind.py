from enum import IntFlag
from ooodev.utils.kind import kind_helper


class HorzVertKind(IntFlag):
    """
    Specifies the Horizontal, Vertical option.

    Int Flags Enum.
    """

    NONE = 0
    """No Option"""
    HORIZONTAL = 1
    """Horizontal Option"""
    VERTICAL = 2
    """Vertical Option"""
    BOTH = HORIZONTAL | VERTICAL
    """Both Options"""

    @staticmethod
    def from_str(s: str) -> "HorzVertKind":
        """
        Gets an ``HorzVertKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``HorzVertKind`` instance.

        Returns:
            HorzVertKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, HorzVertKind)
