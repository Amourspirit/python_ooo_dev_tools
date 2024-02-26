from enum import IntEnum
from ooodev.utils.kind import kind_helper


class DrawingSlideShowKind(IntEnum):
    """DrawPage slide show change constants"""

    AUTO_CHANGE = 1
    """Everything (page change, animation effects) is automatic"""
    CLICK_ALL_CHANGE = 0
    """A mouse-click triggers the next animation effect or page change"""
    CLICK_PAGE_CHANGE = 2
    """Animation effects run automatically, but the user must click on the page to change it"""

    @staticmethod
    def from_str(s: str) -> "DrawingSlideShowKind":
        """
        Gets an ``DrawingSlideShowKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingSlideShowKind`` instance.

        Returns:
            DrawingSlideShowKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingSlideShowKind)
