from __future__ import annotations

from typing import Any, TYPE_CHECKING, TypeVar, Generic, Tuple

from ooodev.adapter.chart2.chart_type_comp import ChartTypeComp
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .coordinate.coordinate_general import CoordinateGeneral
    from .chart_data_series import ChartDataSeries
else:
    CoordinateGeneral = Any

_T = TypeVar("_T", bound="CoordinateGeneral")


class ChartType(
    Generic[_T],
    LoInstPropsPartial,
    ChartTypeComp,
    PropPartial,
    QiPartial,
    ServicePartial,
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
        ChartTypeComp.__init__(self, component=component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        self._owner = owner
        self.get_data_series()

    # region DataSeriesContainerPartial overrides
    def get_data_series(self) -> Tuple[ChartDataSeries[ChartType], ...]:
        """
        retrieve all data series
        """
        from .chart_data_series import ChartDataSeries

        d_series = super().get_data_series()
        if not d_series:
            return ()
        return tuple(ChartDataSeries(owner=self, component=ds) for ds in d_series)

    # endregion DataSeriesContainerPartial overrides

    @property
    def coordinate_sys(self) -> _T:
        """Chart Document"""
        return self._owner
