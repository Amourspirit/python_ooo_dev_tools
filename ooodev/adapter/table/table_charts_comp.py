from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.table.table_charts_partial import TableChartsPartial

if TYPE_CHECKING:
    from com.sun.star.table import TableCharts  # service
    from com.sun.star.table import XTableRows


class TableChartsComp(ComponentBase, TableChartsPartial, EnumerationAccessPartial, IndexAccessPartial):
    """
    Class for managing TableCharts Component.

    Provides methods to access rows via index and to insert and remove rows.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableRows) -> None:
        """
        Constructor

        Args:
            component (XTableRows): UNO Component that implements ``com.sun.star.table.TableCharts`` service.
        """
        ComponentBase.__init__(self, component)
        TableChartsPartial.__init__(self, component=self.component, interface=None)
        EnumerationAccessPartial.__init__(self, component=self.component, interface=None)
        IndexAccessPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableCharts",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TableCharts:
        """TableCharts Component"""
        # pylint: disable=no-member
        return cast("TableCharts", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
