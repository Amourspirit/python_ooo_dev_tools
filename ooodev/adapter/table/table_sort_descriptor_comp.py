from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.table import TableSortDescriptor  # service


class TableSortDescriptorComp(ComponentBase, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing table TableSortDescriptor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TableSortDescriptor) -> None:
        """
        Constructor

        Args:
            component (TableSortDescriptor): UNO table TableSortDescriptor Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableSortDescriptor",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TableSortDescriptor:
        """TableSortDescriptor Component"""
        return cast("TableSortDescriptor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
