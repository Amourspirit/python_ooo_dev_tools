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
from ooodev.format.inner.partial.chart2.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_data_series import ChartDataSeries
    from .chart_doc import ChartDoc


class ChartDataPoint(
    LoInstPropsPartial,
    PropertySetComp,
    ChartDocPropPartial,
    FontEffectsPartial,
    FontOnlyPartial,
    DataPointPropertiesPartial,
    FillPropertiesPartial,
    QiPartial,
    ServicePartial,
    NumbersNumbersPartial,
    FillColorPartial,
    ChartFillGradientPartial,
    ChartFillHatchPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
):
    """
    Class for managing Chart2 Chart Data Point Component.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, owner: ChartDataSeries, chart_doc: ChartDoc, component: Any, lo_inst: LoInst | None = None
    ) -> None:
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
        ChartDocPropPartial.__init__(self, chart_doc=chart_doc)
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
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.chart2.axis.numbers.numbers", component=component, lo_inst=lo_inst
        )
        FillColorPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.area", component=component, lo_inst=lo_inst
        )
        ChartFillGradientPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.area", component=component, lo_inst=lo_inst
        )
        ChartFillHatchPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.area", component=component, lo_inst=lo_inst
        )
        ChartFillImgPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.area", component=component, lo_inst=lo_inst
        )
        ChartFillPatternPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.area", component=component, lo_inst=lo_inst
        )
        self._owner = owner

    @property
    def owner(self) -> ChartDataSeries:
        """Gets the owner of this Chart Data Point."""
        return self._owner
