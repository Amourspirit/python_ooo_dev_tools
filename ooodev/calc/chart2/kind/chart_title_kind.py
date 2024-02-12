from enum import Enum
from ooodev.utils.kind import kind_helper


class ChartTitleKind(Enum):
    """
    Chart Title Kind.
    """

    UNKNOWN = 0
    TITLE = 1
    SUBTITLE = 2
    X_AXIS_TITLE = 3
    X2_AXIS_TITLE = 4
    Y_AXIS_TITLE = 5
    Y2_AXIS_TITLE = 6
    Z_AXIS_TITLE = 7
    Z2_AXIS_TITLE = 8

    @staticmethod
    def from_str(s: str) -> "ChartTitleKind":
        """
        Gets an ``ChartTitleKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``ChartTitleKind`` instance.

        Returns:
            ChartTitleKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ChartTitleKind)
