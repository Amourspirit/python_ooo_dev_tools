from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from .data_source_partial import DataSourcePartial

if TYPE_CHECKING:
    from com.sun.star.chart2.data import DataSource  # service


class DataSourceComp(ComponentBase, DataSourcePartial):
    """
    Class for managing Chart2 Data DataSource Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DataSource) -> None:
        """
        Constructor

        Args:
            component (DataSource): UNO component that implements ``com.sun.star.chart2.data.DataSource`` service.
        """
        ComponentBase.__init__(self, component)
        DataSourcePartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by DataSourcePartial
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> DataSource:
        """DataSource Component"""
        # pylint: disable=no-member
        return cast("DataSource", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
