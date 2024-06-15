from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from ooodev.loader import lo as mLo
from ooodev.utils.comp.prop import Prop
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial
from ooodev.format.inner.partial.chart2.borders.border_line_properties_partial import BorderLinePropertiesPartial
from ooodev.format.inner.partial.area.transparency.transparency_partial import (
    TransparencyPartial as TransparencyTransparency,
)
from ooodev.format.inner.partial.area.transparency.gradient_partial import GradientPartial as TransparencyGradient
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_diagram import ChartDiagram


class ChartFloor(
    Prop["ChartFloor"],
    TheDictionaryPartial,
    ChartDocPropPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    FillPropertiesPartial,
    LinePropertiesPartial,
    FillColorPartial,
    ChartFillGradientPartial,
    ChartFillHatchPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
    BorderLinePropertiesPartial,
    TransparencyTransparency,
    TransparencyGradient,
):
    """
    Class for managing Chart2 Wall.
    """

    def __init__(self, owner: ChartDiagram, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Wall Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        Prop.__init__(self, owner=self, component=component, lo_inst=lo_inst)
        TheDictionaryPartial.__init__(self)
        ChartDocPropPartial.__init__(self, chart_doc=owner.chart_doc)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        FillPropertiesPartial.__init__(self, component=component)
        LinePropertiesPartial.__init__(self, component=component)
        FillColorPartial.__init__(self, factory_name="ooodev.char2.wall.area", component=component, lo_inst=lo_inst)
        ChartFillGradientPartial.__init__(
            self, factory_name="ooodev.char2.wall.area", component=component, lo_inst=lo_inst
        )
        ChartFillHatchPartial.__init__(
            self, factory_name="ooodev.char2.wall.area", component=component, lo_inst=lo_inst
        )
        ChartFillImgPartial.__init__(self, factory_name="ooodev.char2.wall.area", component=component, lo_inst=lo_inst)
        ChartFillPatternPartial.__init__(
            self, factory_name="ooodev.char2.wall.area", component=component, lo_inst=lo_inst
        )
        BorderLinePropertiesPartial.__init__(
            self, factory_name="ooodev.chart2.wall.borders", component=component, lo_inst=lo_inst
        )
        TransparencyTransparency.__init__(
            self, factory_name="ooodev.chart2.wall", component=component, lo_inst=lo_inst
        )
        TransparencyGradient.__init__(
            self, factory_name="ooodev.chart2.wall.transparency", component=component, lo_inst=lo_inst
        )

    # region GradientPartial Overrides

    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return self.chart_doc.component

    # endregion GradientPartial Overrides
