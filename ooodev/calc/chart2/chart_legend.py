from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.chart2.legend_comp import LegendComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import Legend  # service
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_diagram import ChartDiagram
    from .chart_doc import ChartDoc


class ChartLegend(
    LoInstPropsPartial,
    LegendComp,
    FontEffectsPartial,
    FontOnlyPartial,
    EventsPartial,
    QiPartial,
    ServicePartial,
    PropPartial,
):
    """
    Class for managing Chart2 Legend Component.
    """

    def __init__(
        self, owner: ChartDiagram, chart_doc: ChartDoc, component: Legend | None = None, lo_inst: LoInst | None = None
    ) -> None:
        """
        Constructor

        Args:
            component (Legend): UNO Chart2 Legend Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        LegendComp.__init__(self, lo_inst=self.lo_inst, component=component)
        FontEffectsPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        FontOnlyPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        PropPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        QiPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        self._owner = owner
        self._chart_doc = chart_doc

    # region Properties
    @property
    def owner(self) -> ChartDiagram:
        """Owner Chart Diagram"""
        return self._owner

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document."""
        return self._chart_doc

    # endregion Properties
