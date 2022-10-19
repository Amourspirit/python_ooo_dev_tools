from enum import Enum


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
