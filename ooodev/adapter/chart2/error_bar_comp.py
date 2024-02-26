from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet

from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.line_properties_comp import LinePropertiesComp


if TYPE_CHECKING:
    from com.sun.star.chart2 import ErrorBar  # service
    from ooodev.loader.inst.lo_inst import LoInst


class ErrorBarComp(
    LinePropertiesComp,
    PropertySetPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing Chart2 ErrorBar Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, lo_inst: LoInst, component: XPropertySet | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (ErrorBar, optional): UNO Chart2 ErrorBar Component.
        """
        if component is None:
            component = lo_inst.create_instance_mcf(XPropertySet, "com.sun.star.chart2.ErrorBar", raise_err=True)
        LinePropertiesComp.__init__(self, component)  # type: ignore
        PropertySetPartial.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.ErrorBar",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ErrorBar:
        """ErrorBar Component"""
        # pylint: disable=no-member
        return cast("ErrorBar", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
