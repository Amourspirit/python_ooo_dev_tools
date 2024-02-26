from enum import IntEnum
from ooodev.utils.kind import kind_helper


class DataPointLabelPlacementKind(IntEnum):
    """Data Point Label Placement Kind"""

    ABOVE = 0
    CENTER = 1
    LEFT = 4
    BELOW = 6
    RIGHT = 8

    @staticmethod
    def from_str(s: str) -> "DataPointLabelPlacementKind":
        """
        Gets an ``DataPointLabelPlacementKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DataPointLabelPlacementKind`` instance.

        Returns:
            DataPointLabelPlacementKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DataPointLabelPlacementKind)
