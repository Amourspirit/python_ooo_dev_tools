from enum import IntEnum
from . import kind_helper


class DataPointLabelTypeKind(IntEnum):
    """Data Point Label Type Kind"""

    NUMBER = 0
    PERCENT = 1
    CATEGORY = 2
    SYMBOL = 3
    NONE = 4

    @staticmethod
    def from_str(s: str) -> "DataPointLabelTypeKind":
        """
        Gets an ``DataPointLabelTypeKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DataPointLabelTypeKind`` instance.

        Returns:
            DataPointLabelTypeKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DataPointLabelTypeKind)
