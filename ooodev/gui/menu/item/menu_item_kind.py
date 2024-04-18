from __future__ import annotations
from enum import IntEnum
from ooodev.utils.kind import kind_helper


class MenuItemKind(IntEnum):

    SEP = 1
    """Menu Separator"""
    ITEM = 2
    """Menu Item"""
    ITEM_SUBMENU = 3
    """Menu Item with Submenu"""

    @staticmethod
    def from_str(s: str) -> "MenuItemKind":
        """
        Gets an ``MenuItemKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``MenuItemKind`` instance.

        Returns:
            MenuItemKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, MenuItemKind)
