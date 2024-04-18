from enum import IntFlag
from ooodev.utils.kind import kind_helper

# See: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1MenuItemStyle.html

# The const group does not include the NONE value.


class MenuItemStyleKind(IntFlag):
    """
    Flags Enum of Const Group ``com.sun.star.awt.MenuItemStyle``.

    These values are used to specify the properties of a menu item.
    """

    NONE = 0
    """
    No style.
    """
    CHECKABLE = 1
    """
    Specifies an item which can be checked independently.
    """
    RADIOCHECK = 2
    """
    Specifies an item which can be checked dependent on the neighboring items.
    """
    AUTOCHECK = 4
    """
    specifies to check this item automatically on select.
    """

    @staticmethod
    def from_str(s: str) -> "MenuItemStyleKind":
        """
        Gets an ``MenuItemStyleKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``MenuItemStyleKind`` instance.

        Returns:
            MenuItemStyleKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, MenuItemStyleKind)
