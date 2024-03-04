from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart2 import XRegressionCurve

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from com.sun.star.chart2 import XRegressionCurveCalculator
    from ooodev.utils.type_var import UnoInterface


class RegressionCurvePartial:
    """
    Partial class for XRegressionCurve.
    """

    def __init__(self, component: XRegressionCurve, interface: UnoInterface | None = XRegressionCurve) -> None:
        """
        Constructor

        Args:
            component (XRegressionCurve): UNO Component that implements ``com.sun.star.chart2.XRegressionCurve`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XRegressionCurve``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XRegressionCurve
    def get_calculator(self) -> XRegressionCurveCalculator:
        """
        Gets the calculator for the regression curve.
        """
        return self.__component.getCalculator()

    def get_equation_properties(self) -> XPropertySet:
        """
        Gets the properties of the equation.
        """
        return self.__component.getEquationProperties()

    def set_equation_properties(self, x_equation_props: XPropertySet) -> None:
        """
        Sets the properties of the equation.
        """
        self.__component.setEquationProperties(x_equation_props)

    # endregion XRegressionCurve
