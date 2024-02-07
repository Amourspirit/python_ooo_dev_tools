from __future__ import annotations
from typing import Any, TYPE_CHECKING, TypeVar, Generic

from ooodev.adapter.chart2.data_series_comp import DataSeriesComp
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.format.inner.partial.font_effects_partial import FontEffectsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_data_point import ChartDataPoint

_T = TypeVar("_T", bound="ComponentT")


class ChartDataSeries(
    Generic[_T],
    LoInstPropsPartial,
    DataSeriesComp,
    PropPartial,
    QiPartial,
    ServicePartial,
    StylePartial,
    FontEffectsPartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: _T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DataSeriesComp.__init__(self, component=component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        self._owner = owner

    def __getitem__(self, index: int) -> ChartDataPoint:
        return self.get_data_point_by_index(index)

    # region DataSeriesPartial Overrides

    def get_data_point_by_index(self, idx: int) -> ChartDataPoint:
        """

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        from .chart_data_point import ChartDataPoint

        dp = super().get_data_point_by_index(idx)
        return ChartDataPoint(owner=self, component=dp, lo_inst=self.lo_inst)

    # endregion DataSeriesPartial Overrides

    @property
    def owner(self) -> _T:
        """Chart Document"""
        return self._owner
