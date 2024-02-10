from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.chart2.legend_comp import LegendComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.area.transparency.transparency_partial import (
    TransparencyPartial as TransparencyTransparency,
)
from ooodev.format.inner.partial.area.transparency.gradient_partial import GradientPartial as TransparencyGradient
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial


if TYPE_CHECKING:
    from com.sun.star.chart2 import Legend  # service
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_diagram import ChartDiagram
    from .chart_doc import ChartDoc


class ChartLegend(
    LoInstPropsPartial,
    LegendComp,
    ChartDocPropPartial,
    EventsPartial,
    QiPartial,
    ServicePartial,
    PropPartial,
    FontEffectsPartial,
    FontOnlyPartial,
    FillColorPartial,
    TransparencyTransparency,
    TransparencyGradient,
    ChartFillGradientPartial,
    ChartFillHatchPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
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
        ChartDocPropPartial.__init__(self, chart_doc=chart_doc)
        PropPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        QiPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        FontEffectsPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        FontOnlyPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        FillColorPartial.__init__(
            self, factory_name="ooodev.char2.legend.area.color", component=component, lo_inst=lo_inst
        )
        TransparencyTransparency.__init__(
            self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst
        )
        TransparencyGradient.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        ChartFillGradientPartial.__init__(
            self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst
        )
        ChartFillHatchPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        ChartFillImgPartial.__init__(self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst)
        ChartFillPatternPartial.__init__(
            self, factory_name="ooodev.chart2.legend", component=component, lo_inst=lo_inst
        )
        self._owner = owner

    # region GradientPartial Overrides

    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return self.chart_doc.component

    # endregion GradientPartial Overrides

    # region Properties
    @property
    def owner(self) -> ChartDiagram:
        """Owner Chart Diagram"""
        return self._owner

    # endregion Properties
