from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.chart2 import XScaling
from ooodev.adapter.chart2.axis_comp import AxisComp
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.kind.curve_kind import CurveKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.partial.borders.chart2.axis_line_properties_partial import AxisLinePropertiesPartial
from ooodev.format.inner.partial.chart2.axis.positioning.chart2_axis_pos_axis_line_partial import (
    Chart2AxisPosAxisLinePartial,
)
from ooodev.format.inner.partial.chart2.axis.positioning.chart2_axis_pos_interval_marks_partial import (
    Chart2AxisPosIntervalMarksPartial,
)
from ooodev.format.inner.partial.chart2.axis.positioning.chart2_axis_pos_label_position_partial import (
    Chart2AxisPosLabelPositionPartial,
)
from ooodev.format.inner.partial.chart2.axis.positioning.chart2_axis_pos_position_axis_partial import (
    Chart2AxisPosPositionAxisPartial,
)
from ooodev.format.inner.partial.chart2.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.format.inner.partial.chart2.grid.chart2_grid_line_partial import Chart2GridLinePartial

if TYPE_CHECKING:
    from .chart_doc import ChartDoc
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from .chart_title import ChartTitle


class ChartAxis(
    LoInstPropsPartial,
    AxisComp,
    EventsPartial,
    ChartDocPropPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    StylePartial,
    FontOnlyPartial,
    FontEffectsPartial,
    AxisLinePropertiesPartial,
    Chart2AxisPosAxisLinePartial,
    Chart2AxisPosIntervalMarksPartial,
    Chart2AxisPosLabelPositionPartial,
    Chart2AxisPosPositionAxisPartial,
    NumbersNumbersPartial,
    Chart2GridLinePartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: ChartDoc, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        AxisComp.__init__(self, component=component)
        EventsPartial.__init__(self)
        ChartDocPropPartial.__init__(self, chart_doc=owner)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        FontEffectsPartial.__init__(self, factory_name="ooodev.chart2.axis", component=component, lo_inst=lo_inst)
        FontOnlyPartial.__init__(self, factory_name="ooodev.chart2.axis", component=component, lo_inst=lo_inst)
        AxisLinePropertiesPartial.__init__(
            self, factory_name="ooodev.chart2.axis.line", component=component, lo_inst=lo_inst
        )
        Chart2AxisPosAxisLinePartial.__init__(
            self, factory_name="ooodev.chart2.axis.pos.line", component=component, lo_inst=lo_inst
        )
        Chart2AxisPosIntervalMarksPartial.__init__(
            self, factory_name="ooodev.chart2.axis.pos.interval_marks", component=component, lo_inst=lo_inst
        )
        Chart2AxisPosLabelPositionPartial.__init__(
            self, factory_name="ooodev.chart2.axis.pos.position", component=component, lo_inst=lo_inst
        )
        Chart2AxisPosPositionAxisPartial.__init__(
            self, factory_name="ooodev.chart2.axis.line", component=component, lo_inst=lo_inst
        )
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.chart2.axis.numbers.numbers", component=component, lo_inst=lo_inst
        )
        grid_props = self.get_grid_properties()
        Chart2GridLinePartial.__init__(
            self, factory_name="ooodev.chart2.grid.line", component=grid_props, lo_inst=lo_inst
        )

    # region StylePartial Overrides

    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        mChart2.Chart2._style_title(self.__component, styles)

    # endregion

    def get_title(self) -> ChartTitle[ChartAxis] | None:
        """
        Gets the Chart Title Component.

        Raises:
            ChartError: If error occurs.

        Returns:
            ChartTitle: Chart Title Component.
        """
        from com.sun.star.chart2 import XTitled
        from .chart_title import ChartTitle

        try:
            titled = self.qi(XTitled, True)
            comp = titled.getTitleObject()
            if comp is None:
                return None
            return ChartTitle(owner=self, chart_doc=self.chart_doc, component=comp, lo_inst=self.lo_inst)
        except Exception as e:
            raise mEx.ChartError("Error getting axis title") from e

    def set_title(self, title: str) -> ChartTitle:
        """
        Sets Chart Title.

        Args:
            title (str): Title text.

        Raises:
            ChartError: If error occurs.

        Returns:
            ChartTitle: Chart Title Component.
        """
        from com.sun.star.chart2 import XTitled
        from com.sun.star.chart2 import XTitle
        from com.sun.star.chart2 import XFormattedString
        from .chart_title import ChartTitle

        try:
            x_title = self.lo_inst.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
            x_title_str = self.lo_inst.create_instance_mcf(
                XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
            )
            x_title_str.setString(title)

            title_arr = (x_title_str,)
            x_title.setText(title_arr)

            titled = self.qi(XTitled, True)
            titled.setTitleObject(x_title)
            return ChartTitle(
                owner=self, chart_doc=self.chart_doc, component=titled.getTitleObject(), lo_inst=self.lo_inst
            )
        except Exception as e:
            raise mEx.ChartError("Error setting axis title") from e

    def scale(self, scale_type: CurveKind) -> None:
        """
        Scales the axis.

        Args:
            scale_type (CurveKind): Scale kind

        Raises:
            ChartError: If error occurs.

        Returns:
            None:

        Note:
            Supported types of ``scale_type`` are ``LINEAR``, ``LOGARITHMIC``, ``EXPONENTIAL`` and ``POWER``.
            If ``scale_type``  is not supported then the ``Scaling`` is not set.
        """
        try:
            sd = self.get_scale_data()
            s = None
            if scale_type == CurveKind.LINEAR:
                s = "LinearScaling"
            elif scale_type == CurveKind.LOGARITHMIC:
                s = "LogarithmicScaling"
            elif scale_type == CurveKind.EXPONENTIAL:
                s = "ExponentialScaling"
            elif scale_type == CurveKind.POWER:
                s = "PowerScaling"
            if s is None:
                self.lo_inst.print(f'Did not recognize scaling type: "{scale_type}"')
            else:
                sd.Scaling = self.lo_inst.create_instance_mcf(XScaling, f"com.sun.star.chart2.{s}", raise_err=True)
            self.set_scale_data(sd)
        except Exception as e:
            raise mEx.ChartError("Error setting axis scale") from e
