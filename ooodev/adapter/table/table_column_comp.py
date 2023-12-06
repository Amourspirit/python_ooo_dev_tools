from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.table import TableColumn  # service


class TableColumnComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing table TableColumn Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TableColumn) -> None:
        """
        Constructor

        Args:
            component (TableColumn): UNO table TableColumn Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableColumn",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TableColumn:
        """TableColumn Component"""
        return cast("TableColumn", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
