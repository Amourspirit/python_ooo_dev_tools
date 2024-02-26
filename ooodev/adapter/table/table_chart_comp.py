from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.document.embedded_object_supplier_partial import EmbeddedObjectSupplierPartial
from ooodev.adapter.table.table_chart_partial import TableChartPartial

if TYPE_CHECKING:
    from com.sun.star.table import TableChart  # service
    from com.sun.star.table import XTableRows


class TableChartComp(ComponentBase, TableChartPartial, EmbeddedObjectSupplierPartial):
    """
    Class for managing TableChart Component.

    Provides methods to access rows via index and to insert and remove rows.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableRows) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.table.TableChart`` service.
        """
        ComponentBase.__init__(self, component)
        TableChartPartial.__init__(self, component=self.component, interface=None)
        EmbeddedObjectSupplierPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableChart",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TableChart:
        """TableChart Component"""
        # pylint: disable=no-member
        return cast("TableChart", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
