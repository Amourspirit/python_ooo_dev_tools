from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.table.cell_range_partial import CellRangePartial


if TYPE_CHECKING:
    from com.sun.star.table import CellRange  # service


class CellRangeComp(ComponentBase, CellRangePartial, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing table CellRange Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellRange) -> None:
        """
        Constructor

        Args:
            component (CellRange): UNO table CellRange Component.
        """
        ComponentBase.__init__(self, component)
        CellRangePartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.CellRange",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CellRange:
        """CellRange Component"""
        # pylint: disable=no-member
        return cast("CellRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
