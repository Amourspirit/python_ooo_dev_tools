from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.chart2 import XDataSeries
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.chart2.data_point_properties_partial import DataPointPropertiesPartial
from ooodev.adapter.chart2.data_series_partial import DataSeriesPartial
from ooodev.adapter.chart2.data.data_sink_partial import DataSinkPartial
from ooodev.adapter.chart2.data.data_source_partial import DataSourcePartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import DataSeries  # service
    from ooodev.loader.inst.lo_inst import LoInst


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

    def __init__(self, lo_inst: LoInst, component: DataSeries | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (DataSeries, optional): UNO Chart2 DataSeries Component.
        """
        if component is None:
            component = cast(
                "DataSeries",
                lo_inst.create_instance_mcf(XDataSeries, "com.sun.star.chart2.DataSeries", raise_err=True),
            )
        ComponentBase.__init__(self, component)
        DataPointPropertiesPartial.__init__(self, component=component)
        DataSeriesPartial.__init__(self, component=component, interface=None)  # type: ignore
        DataSinkPartial.__init__(self, component=component, interface=None)
        DataSourcePartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
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
        # pylint: disable=no-member
        return cast("DataSeries", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
