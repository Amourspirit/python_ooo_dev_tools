from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.chart2.legend_comp import LegendComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import Legend  # service
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_diagram import ChartDiagram


class ChartLegend(LoInstPropsPartial, LegendComp, EventsPartial, QiPartial, ServicePartial, PropPartial):
    """
    Class for managing Chart2 Legend Component.
    """

    def __init__(self, owner: ChartDiagram, component: Legend | None = None, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Legend): UNO Chart2 Legend Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        LegendComp.__init__(self, lo_inst=self.lo_inst, component=component)
        PropPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        QiPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        self._owner = owner

    # region Properties
    @property
    def owner(self) -> ChartDiagram:
        """Owner Chart Diagram"""
        return self._owner

    # endregion Properties
