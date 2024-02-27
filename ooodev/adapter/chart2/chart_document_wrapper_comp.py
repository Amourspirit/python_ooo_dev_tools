from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.chart2 import ChartDocumentWrapper  # service


class ChartDocumentWrapperComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Chart2 ChartDocumentWrapper Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ChartDocumentWrapper) -> None:
        """
        Constructor

        Args:
            component (ChartDocumentWrapper): UNO Chart2 ChartDocumentWrapper Component.
        """
        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.ChartDocumentWrapper",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ChartDocumentWrapper:
        """ChartDocumentWrapper Component"""
        # pylint: disable=no-member
        return cast("ChartDocumentWrapper", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
