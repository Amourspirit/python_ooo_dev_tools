from enum import Enum
from . import kind_helper


class DataRoleKind(str, Enum):
    """
    Represents DataSequenceRole

    See Also:
        `DataSequenceRole API <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2_1_1data.html#a340775895b509a7f80ba895767123429>`_
    """

    CATEGORIES = "categories"
    """Values are used for categories in the diagram"""
    ERROR_BARS_X_NEGATIVE = "error-bars-x-negative"
    """Values are used as error-information in negative x-direction for displaying error-bars"""
    ERROR_BARS_X_POSITIVE = "error-bars-x-positive"
    """Values are used as error-information in positive x-direction for displaying error-bars"""
    ERROR_BARS_Y_NEGATIVE = "error-bars-y-negative"
    """Values are used as error-information in negative y-direction for displaying error-bars"""
    ERROR_BARS_Y_POSITIVE = "error-bars-y-positive"
    """Values are used as error-information in positive y-direction for displaying error-bars"""
    LABEL = "label"
    """Values are used as a label for a series. Usually, you will have just one cell containing a string."""
    SIZES = "sizes"
    """values are used as radius of the bubbles in a Bubble-Diagram."""
    VALUES_FIRST = "values-first"
    """``Candle-stick`` chart - the first value of a series of values. In a stock-chart this would be the opening course."""
    VALUES_LAST = "values-last"
    """``Candle-stick`` chart - the last value of a series of values. In a stock-chart this would be the closing course."""
    VALUES_MAX = "values-max"
    """``Candle-stick`` chart - the maximum value of a series of values. In a stock-chart this would be the highest course that occurred during trading."""
    VALUES_MIN = "values-min"
    """``Candle-stick`` chart - the minimum value of a series of values. In a stock-chart this would be the lowest course that occurred during trading."""
    VALUES_X = "values-x"
    """Values are used as x-values in an ``XY`` - or bubble diagram."""
    VALUES_Y = "values-y"
    """Values are used as y-values in an ``XY-Diagram`` or as values in a bar, line, etc. chart."""
    VALUES_Z = "values-z"
    """Values may be used as z-values in a three-dimensional ``XYZ-Diagram`` or a surface-chart."""

    def __str__(self) -> str:
        return str(self.value)

    @staticmethod
    def from_str(s: str) -> "DataRoleKind":
        """
        Gets an ``DataRoleKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DataRoleKind`` instance.

        Returns:
            DataRoleKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DataRoleKind)
