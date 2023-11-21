from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.table import Cell  # service


class CellComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing table Cell Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Cell) -> None:
        """
        Constructor

        Args:
            component (Cell): UNO table Cell Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._get_generic_args()
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table .Cell",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Cell:
        """Cell Component"""
        return cast("Cell", self._get_component())

    # endregion Properties
