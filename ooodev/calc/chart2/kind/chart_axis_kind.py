from enum import Enum
from ooodev.utils.kind import kind_helper


class ChartAxisKind(Enum):
    # these values line up with ChartTitleKind
    X = 3
    X2 = 4
    Y = 5
    Y2 = 6
    Z = 7
    Z2 = 8

    @staticmethod
    def from_str(s: str) -> "ChartAxisKind":
        """
        Gets an ``ChartAxisKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``ChartAxisKind`` instance.

        Returns:
            ChartAxisKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ChartAxisKind)
