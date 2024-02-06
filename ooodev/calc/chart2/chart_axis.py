from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.adapter.chart2.axis_comp import AxisComp
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial

if TYPE_CHECKING:
    from .chart_doc import ChartDoc
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from .chart_title import ChartTitle


class ChartAxis(LoInstPropsPartial, AxisComp, PropPartial, QiPartial, ServicePartial, StylePartial):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: ChartDoc, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        AxisComp.__init__(self, component=component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        self._owner = owner

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
        """Gets the Chart Title Component."""
        from com.sun.star.chart2 import XTitled
        from .chart_title import ChartTitle

        titled = self.qi(XTitled, True)
        comp = titled.getTitleObject()
        if comp is None:
            return None
        return ChartTitle(owner=self, component=comp, lo_inst=self.lo_inst)

    def set_title(self, title: str) -> ChartTitle:
        """Sets Chart Title."""
        from com.sun.star.chart2 import XTitled
        from com.sun.star.chart2 import XTitle
        from com.sun.star.chart2 import XFormattedString
        from .chart_title import ChartTitle

        x_title = self.lo_inst.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
        x_title_str = self.lo_inst.create_instance_mcf(
            XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
        )
        x_title_str.setString(title)

        title_arr = (x_title_str,)
        x_title.setText(title_arr)

        titled = self.qi(XTitled, True)
        titled.setTitleObject(x_title)
        return ChartTitle(owner=self, component=titled.getTitleObject(), lo_inst=self.lo_inst)

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document"""
        return self._owner
