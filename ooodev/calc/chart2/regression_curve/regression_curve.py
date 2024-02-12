from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.chart2.regression_curve_partial import RegressionCurvePartial
from ooodev.adapter.component_base import ComponentBase
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils import info as mInfo
from ooodev.utils.comp.prop import Prop
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_doc import ChartDoc
    from com.sun.star.chart2 import RegressionCurve as UnoRegressionCurve


class RegressionCurve(
    LoInstPropsPartial,
    ComponentBase,
    ChartDocPropPartial,
    RegressionCurvePartial,
    PropertySetPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
):

    def __init__(self, owner: ChartDoc, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ComponentBase.__init__(self, component=component)
        ChartDocPropPartial.__init__(self, chart_doc=owner)
        RegressionCurvePartial.__init__(self, component=component, interface=None)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        PropertySetPartial.__init__(self, component=component, interface=None)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return (
            "com.sun.star.chart2.LinearRegressionCurve",
            "com.sun.star.chart2.LogarithmicRegressionCurve",
            "com.sun.star.chart2.ExponentialRegressionCurve",
            "com.sun.star.chart2.PotentialRegressionCurve",
            "com.sun.star.chart2.PolynomialRegressionCurve",
            "com.sun.star.chart2.MovingAverageRegressionCurve",
        )

    # region RegressionCurvePartial overrides
    def get_equation_properties(self) -> Prop[RegressionCurve]:
        """
        Gets the properties of the equation.
        """
        ps = super().get_equation_properties()
        return Prop(owner=self, component=ps, lo_inst=self.lo_inst)  # type: ignore

    def set_equation_properties(self, x_equation_props: XPropertySet | Prop) -> None:
        """
        Sets the properties of the equation.
        """
        if mInfo.Info.is_instance(x_equation_props, Prop):
            self.__component.setEquationProperties(x_equation_props.component)
            return
        self.__component.setEquationProperties(x_equation_props)

    # endregion RegressionCurvePartial overrides
    # endregion Overrides

    def eval_curve(self) -> None:
        """
        Uses ``XRegressionCurve.getCalculator()`` to access the ``XRegressionCurveCalculator`` interface.
        It sets up the data and parameters for a particular curve, and prints the results of curve fitting to the console.

        Returns:
            None:
        """
        mChart2.Chart2.eval_curve(self.chart_doc.component, self.component)

    @property
    def component(self) -> UnoRegressionCurve:
        """RegressionCurve Component"""
        return cast("UnoRegressionCurve", self._ComponentBase__get_component())  # type: ignore
