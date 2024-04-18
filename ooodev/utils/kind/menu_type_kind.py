from enum import IntEnum
from ooodev.utils.kind import kind_helper


class MenuTypeKind(IntEnum):
    """
    Specifies the orientation of the control.

    Similar to ``com.sun.star.awt.MenuItemType`` but the value are the integer representation of the enum.
    """

    DONTKNOW = 0
    """Specifies that the menu item type is unknown."""
    STRING = 1
    """Specifies that the menu item has a text."""
    IMAGE = 2
    """Specifies that the menu item has a text and an image."""
    STRINGIMAGE = 3
    """Specifies that the menu item has a text and an image."""
    SEPARATOR = 4
    """Specifies that the menu item is a separator."""

    @staticmethod
    def from_str(s: str) -> "MenuTypeKind":
        """
        Gets an ``MenuTypeKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``MenuTypeKind`` instance.

        Returns:
            MenuTypeKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, MenuTypeKind)
