from __future__ import annotations
from typing import Any, TYPE_CHECKING, Type
from ooodev.loader import lo as mLo
from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.format.inner.partial.font_effects_partial import FontEffectsPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.adapter.chart2.data_point_properties_partial import DataPointPropertiesPartial
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_data_series import ChartDataSeries


class ChartDataPoint(
    LoInstPropsPartial, PropertySetComp, FontEffectsPartial, DataPointPropertiesPartial, FillPropertiesPartial
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
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        PropertySetComp.__init__(self, component)
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        DataPointPropertiesPartial.__init__(self, component=component)
        FillPropertiesPartial.__init__(self, component=component)
        self._owner = owner

    @property
    def owner(self) -> ChartDataSeries:
        """Gets the owner of this Chart Data Point."""
        return self._owner
