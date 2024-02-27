from __future__ import annotations
from typing import Any, TYPE_CHECKING, List, TypeVar, Generic
import uno
from com.sun.star.chart2.data import XDataSource

from ooodev.mock import mock_g
from ooodev.adapter.chart2.data_series_comp import DataSeriesComp
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial
from ooodev.format.inner.partial.chart2.borders.border_line_properties_partial import BorderLinePropertiesPartial
from ooodev.format.inner.partial.chart2.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.kind.data_point_label_type_kind import DataPointLabelTypeKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
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
    from ooodev.calc.chart2.chart_doc import ChartDoc
    from ooodev.calc.chart2.chart_data_point import ChartDataPoint
    from ooodev.calc.chart2.data.data_source import DataSource

_T = TypeVar("_T", bound="ComponentT")


class ChartDataSeries(
    Generic[_T],
    LoInstPropsPartial,
    DataSeriesComp,
    ChartDocPropPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    StylePartial,
    FontEffectsPartial,
    FontOnlyPartial,
    FontPartial,
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
    Chart2DataLabelTextAttributePartial,
    Chart2DataLabelAttribOptPartial,
    Chart2DataLabelPercentFormatPartial,
    Chart2DataLabelOrientationPartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(
        self, owner: _T, chart_doc: ChartDoc, component: Any | None = None, lo_inst: LoInst | None = None
    ) -> None:
        """
        Constructor

        Args:
            component (Any, optional): UNO Chart2 Title Component. If None, it will be created using ``lo_inst``.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DataSeriesComp.__init__(self, lo_inst=self.lo_inst, component=component)
        ChartDocPropPartial.__init__(self, chart_doc=chart_doc)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=chart_doc.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=chart_doc.calc_sheet)
        StylePartial.__init__(self, component=component)
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        FontOnlyPartial.__init__(
            self, factory_name="ooodev.chart2.series.data_labels", component=component, lo_inst=lo_inst
        )
        FontPartial.__init__(self, factory_name="ooodev.general_style.text", component=component, lo_inst=lo_inst)
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.chart2.axis.numbers.numbers", component=component, lo_inst=lo_inst
        )
        FillColorPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.area", component=component, lo_inst=lo_inst
        )
        ChartFillGradientPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.area", component=component, lo_inst=lo_inst
        )
        ChartFillHatchPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.area", component=component, lo_inst=lo_inst
        )
        ChartFillImgPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.area", component=component, lo_inst=lo_inst
        )
        ChartFillPatternPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.area", component=component, lo_inst=lo_inst
        )
        BorderLinePropertiesPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.borders", component=component, lo_inst=lo_inst
        )
        TransparencyTransparency.__init__(
            self, factory_name="ooodev.char2.series.data_series.transparency", component=component, lo_inst=lo_inst
        )
        TransparencyGradient.__init__(
            self, factory_name="ooodev.char2.series.data_series.transparency", component=component, lo_inst=lo_inst
        )
        DataLabelBorderPartial.__init__(
            self, factory_name="ooodev.char2.series.data_series.label.borders", component=component, lo_inst=lo_inst
        )
        Chart2DataLabelTextAttributePartial.__init__(self, component=component)
        Chart2DataLabelAttribOptPartial.__init__(self, component=component)
        Chart2DataLabelPercentFormatPartial.__init__(self, component=component)
        Chart2DataLabelOrientationPartial.__init__(self, component=component)

        self._owner = owner

    def __getitem__(self, key: int) -> ChartDataPoint:
        """
        Gets the data point at the specified index.

        Args:
            key (int): The index. When getting by index can be a negative value to get from the end.

        Returns:
            ChartDataPoint: The sheet with the specified index or name.
        """
        return self.get_data_point_by_index(key)

    def get_data_points(self) -> List[ChartDataPoint]:
        """ "
        Gets all the data points of the series.

        Returns:
            List[ChartDataPoint]: List of data points.
        """
        # pylint: disable=import-outside-toplevel
        from .chart_data_point import ChartDataPoint

        lst = []
        i = 0
        comp = self.component
        while True:
            try:
                props = comp.getDataPointByIndex(i)
                if props is not None:
                    lst.append(
                        ChartDataPoint(owner=self, chart_doc=self.chart_doc, component=props, lo_inst=self.lo_inst)
                    )
                i += 1
            except Exception:
                props = None

            if props is None:
                break
        return lst

    # region DataSeriesPartial Overrides

    def get_data_point_by_index(self, idx: int) -> ChartDataPoint:
        """
        Gets a data point by index.

        Args:
            idx (int): Index of data point. Can be a negative value to index from the end of the list.

        Raises:
            IndexError: If index is out of range.
        """
        if idx < 0:
            points = self.get_data_points()
            count = len(points)
            if count == 0:
                raise IndexError("Index out of range")
            index = mGenUtil.Util.get_index(idx, count, False)
            return points[index]

        # pylint: disable=import-outside-toplevel
        from .chart_data_point import ChartDataPoint

        dp = super().get_data_point_by_index(idx)
        if dp is None:
            raise IndexError("Index out of range")
        return ChartDataPoint(owner=self, chart_doc=self.chart_doc, component=dp, lo_inst=self.lo_inst)

    # endregion DataSeriesPartial Overrides

    # region GradientPartial Overrides

    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return self.chart_doc.component

    # endregion GradientPartial Overrides

    def get_data_source(self) -> DataSource:
        """
        Get data source of a chart for a given chart type.

        Raises:
            ChartError: If any error occurs.

        Returns:
            DataSource: Chart data source
        """
        # pylint: disable=import-outside-toplevel
        from .data.data_source import DataSource

        try:
            src = self.qi(XDataSource, True)
            return DataSource(owner=self, component=src, lo_inst=self.lo_inst)
        except Exception as e:
            raise mEx.ChartError("Error getting data source") from e

    def set_data_point_labels(self, label_type: DataPointLabelTypeKind) -> None:
        """
        Set data point labels for a given chart type.

        Args:
            label_type (DataPointLabelTypeKind): Data point label type.

        Raises:
            ChartError: If any error occurs.

        Returns:
            None:

        Hint:
            - ``DataPointLabelTypeKind`` can be imported from ``ooodev.utils.kind.data_point_label_type_kind``
        """
        try:
            dp_label = self.label

            dp_label.ShowNumber = False
            dp_label.ShowCategoryName = False
            dp_label.ShowLegendSymbol = False
            if label_type == DataPointLabelTypeKind.NUMBER:
                dp_label.ShowNumber = True
            elif label_type == DataPointLabelTypeKind.PERCENT:
                dp_label.ShowNumber = True
                dp_label.ShowNumberInPercent = True
            elif label_type == DataPointLabelTypeKind.CATEGORY:
                dp_label.ShowCategoryName = True
            elif label_type == DataPointLabelTypeKind.SYMBOL:
                dp_label.ShowLegendSymbol = True
            elif label_type != DataPointLabelTypeKind.NONE:
                raise mEx.UnKnownError("label_type is of unknown type")
            self.label = dp_label
        except Exception as e:
            raise mEx.ChartError("Error setting data point labels") from e

    @property
    def owner(self) -> _T:
        """Owner"""
        return self._owner


if mock_g.FULL_IMPORT:
    from .chart_data_point import ChartDataPoint
    from .data.data_source import DataSource
