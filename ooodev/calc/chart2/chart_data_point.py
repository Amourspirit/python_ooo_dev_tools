from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.loader import lo as mLo
from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
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
from ooodev.format.inner.partial.chart2.borders.border_line_properties_partial import BorderLinePropertiesPartial
from ooodev.format.inner.partial.area.transparency.transparency_partial import (
    TransparencyPartial as TransparencyTransparency,
)
from ooodev.format.inner.partial.area.transparency.gradient_partial import GradientPartial as TransparencyGradient
from ooodev.format.inner.partial.chart2.series.data_labels.borders.data_label_border_partial import (
    DataLabelBorderPartial,
)
from ooodev.format.inner.partial.chart2.series.data_labels.data_labels.chart2_data_label_attrib_opt_partial import (
    Chart2DataLabelAttribOptPartial,
)

from ooodev.format.inner.partial.chart2.series.data_labels.data_labels.chart2_data_label_percent_format_partial import (
    Chart2DataLabelPercentFormatPartial,
)
from ooodev.format.inner.partial.chart2.series.data_labels.data_labels.chart2_data_label_orientation_partial import (
    Chart2DataLabelOrientationPartial,
)
from ooodev.format.inner.partial.chart2.series.data_labels.data_labels.chart2_data_label_text_attribute_partial import (
    Chart2DataLabelTextAttributePartial,
)
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_data_series import ChartDataSeries
    from ooodev.calc.chart2.chart_doc import ChartDoc


class ChartDataPoint(
    LoInstPropsPartial,
    PropertySetComp,
    ChartDocPropPartial,
    FontEffectsPartial,
    FontOnlyPartial,
    FontPartial,
    DataPointPropertiesPartial,
    FillPropertiesPartial,
    QiPartial,
    ServicePartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    NumbersNumbersPartial,
    FillColorPartial,
    ChartFillGradientPartial,
    ChartFillHatchPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
    BorderLinePropertiesPartial,
    TransparencyTransparency,
    TransparencyGradient,
    DataLabelBorderPartial,
    Chart2DataLabelAttribOptPartial,
    Chart2DataLabelPercentFormatPartial,
    Chart2DataLabelOrientationPartial,
    Chart2DataLabelTextAttributePartial,
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
        FontPartial.__init__(self, factory_name="ooodev.general_style.text", component=component, lo_inst=lo_inst)
        DataPointPropertiesPartial.__init__(self, component=component)
        FillPropertiesPartial.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=chart_doc.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=chart_doc.calc_sheet)
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
        BorderLinePropertiesPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.borders", component=component, lo_inst=lo_inst
        )
        TransparencyTransparency.__init__(
            self, factory_name="ooodev.char2.series.data_point.transparency", component=component, lo_inst=lo_inst
        )
        TransparencyGradient.__init__(
            self, factory_name="ooodev.char2.series.data_point.transparency", component=component, lo_inst=lo_inst
        )
        DataLabelBorderPartial.__init__(
            self, factory_name="ooodev.char2.series.data_point.label.borders", component=component, lo_inst=lo_inst
        )
        Chart2DataLabelAttribOptPartial.__init__(self, component=component)
        Chart2DataLabelPercentFormatPartial.__init__(self, component=component)
        Chart2DataLabelOrientationPartial.__init__(self, component=component)
        Chart2DataLabelTextAttributePartial.__init__(self, component=component)
        self._owner = owner

    # region GradientPartial Overrides

    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return self.chart_doc.component

    # endregion GradientPartial Overrides

    @property
    def owner(self) -> ChartDataSeries:
        """Gets the owner of this Chart Data Point."""
        return self._owner
