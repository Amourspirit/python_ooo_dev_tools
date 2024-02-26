from enum import IntEnum
from ooodev.utils.kind import kind_helper


class ShapeCombKind(IntEnum):
    """Shape combine Kind"""

    MERGE = 0
    INTERSECT = 1
    SUBTRACT = 2
    COMBINE = 3

    @staticmethod
    def from_str(s: str) -> "ShapeCombKind":
        """
        Gets an ``ShapeCombKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``ShapeCombKind`` instance.

        Returns:
            ShapeCombKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ShapeCombKind)
