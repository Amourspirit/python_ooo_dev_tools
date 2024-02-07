from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase

from .data_point_properties_partial import DataPointPropertiesPartial
from .data_series_partial import DataSeriesPartial
from .data.data_sink_partial import DataSinkPartial
from .data.data_source_partial import DataSourcePartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import DataSeries  # service


class DataSeriesComp(
    ComponentBase,
    DataPointPropertiesPartial,
    DataSeriesPartial,
    DataSinkPartial,
    DataSourcePartial,
    PropertiesChangeImplement,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing Chart2 DataSeries Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DataSeries) -> None:
        """
        Constructor

        Args:
            component (DataSeries): UNO Chart2 DataSeries Component.
        """
        ComponentBase.__init__(self, component)
        DataPointPropertiesPartial.__init__(self, component=component)
        DataSeriesPartial.__init__(self, component=component, interface=None)  # type: ignore
        DataSinkPartial.__init__(self, component=component, interface=None)
        DataSourcePartial.__init__(self, component=component, interface=None)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.DataSeries",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> DataSeries:
        """DataSeries Component"""
        return cast("DataSeries", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
