from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.loader import lo as mLo
from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.adapter.chart2.data_point_properties_partial import DataPointPropertiesPartial
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_data_series import ChartDataSeries


class ChartDataPoint(
    LoInstPropsPartial,
    PropertySetComp,
    FontEffectsPartial,
    FontOnlyPartial,
    DataPointPropertiesPartial,
    FillPropertiesPartial,
    QiPartial,
    ServicePartial,
):
    """
    Class for managing Chart2 Chart Data Point Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, owner: ChartDataSeries, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Chart Data Point Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        PropertySetComp.__init__(self, component)
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        FontOnlyPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        DataPointPropertiesPartial.__init__(self, component=component)
        FillPropertiesPartial.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        self._owner = owner

    @property
    def owner(self) -> ChartDataSeries:
        """Gets the owner of this Chart Data Point."""
        return self._owner
