from enum import Enum
from ooodev.utils.kind import kind_helper


class CurveKind(Enum):
    """
    Curve Kind Enum. Regression line types.

    Usage:

    .. code-block:: python

        # compare enums
        >>> x = CurveKind.MOVING_AVERAGE
        >>> y = CurveKind.MOVING_AVERAGE
        >>> assert x == y

        >>> print(x.curve)
        5
        >>> print(x.label)
        Moving average
        >>> print(x.name)
        MOVING_AVERAGE
        >>> print(CurveKind.LINEAR.to_namespace())
        com.sun.star.chart2.LinearRegressionCurve
    """

    LINEAR = 0, "Linear", "LinearRegressionCurve"
    LOGARITHMIC = 1, "Logarithmic", "LogarithmicRegressionCurve"
    EXPONENTIAL = 2, "Exponential", "ExponentialRegressionCurve"
    POWER = 3, "Power", "PotentialRegressionCurve"
    POLYNOMIAL = 4, "Polynomial", "PolynomialRegressionCurve"
    MOVING_AVERAGE = 5, "Moving average", "MovingAverageRegressionCurve"

    def __init__(self, curve: int, label: str, ns: str):
        self.curve = curve
        self.label = label
        self.ns_regression_curve = ns

    def to_namespace(self) -> str:
        """
        Gets the full UNO namespace value of CurveKind instance.

        Returns:
            str: String namespace such as ``com.sun.star.chart2.LinearRegressionCurve``
        """
        return f"com.sun.star.chart2.{self.ns_regression_curve}"

    @staticmethod
    def from_str(s: str) -> "CurveKind":
        """
        Gets an ``CurveKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``CurveKind`` instance.

        Returns:
            CurveKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, CurveKind)
