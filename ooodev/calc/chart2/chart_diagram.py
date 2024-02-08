from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from ooodev.adapter.chart2.diagram_comp import DiagramComp
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_doc import ChartDoc
    from .coordinate.coordinate_general import CoordinateGeneral
    from .chart_title import ChartTitle
    from .chart_legend import ChartLegend


class ChartDiagram(LoInstPropsPartial, DiagramComp, QiPartial, ServicePartial, StylePartial):
    """
    Class for managing Chart2 Diagram.
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
        DiagramComp.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        self._owner = owner

    def get_title(self) -> ChartTitle[ChartDiagram] | None:
        """Gets the Title Diagram Component. This might be considered to be a subtitle."""
        from .chart_title import ChartTitle

        comp = self.get_title_object()
        if comp is None:
            return None
        return ChartTitle(owner=self, component=comp, lo_inst=self.lo_inst)

    def set_title(self, title: str) -> ChartTitle[ChartDiagram]:
        """Adds a Chart Title."""
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

    def get_coordinate_system(self) -> CoordinateGeneral | None:
        """Gets the first Coordinate System Component."""

        coord_sys = super().get_coordinate_systems()
        if not coord_sys:
            return None
        first = coord_sys[0]
        if mInfo.Info.support_service(first, "com.sun.star.chart2.CoordinateSystem"):
            from .coordinate.coordinate_system import CoordinateSystem

            result = CoordinateSystem(owner=self, component=first, lo_inst=self.lo_inst)
        else:
            from .coordinate.coordinate_general import CoordinateGeneral

            result = CoordinateGeneral(owner=self, component=first, lo_inst=self.lo_inst)

        return result

    # region CoordinateSystemContainerPartial overrides
    def get_coordinate_systems(self) -> Tuple[CoordinateGeneral, ...]:
        """
        Gets all coordinate systems
        """
        from .coordinate.coordinate_system import CoordinateSystem
        from .coordinate.coordinate_general import CoordinateGeneral

        result = []
        coord_sys = super().get_coordinate_systems()
        for sys in coord_sys:
            if mInfo.Info.support_service(sys, "com.sun.star.chart2.CoordinateSystem"):
                result.append(CoordinateSystem(owner=self, component=coord_sys[0], lo_inst=self.lo_inst))
            else:
                result.append(CoordinateGeneral(owner=self, component=coord_sys[0], lo_inst=self.lo_inst))
        return tuple(result)

    # endregion CoordinateSystemContainerPartial overrides

    def get_legend(self) -> ChartLegend | None:
        """
        Gets the Legend Component.

        Returns:
            ChartLegend | None: Legend Component if found, otherwise ``None``.
        """
        legend = super().get_legend()
        if legend is None:
            return None
        from .chart_legend import ChartLegend

        return ChartLegend(owner=self, component=legend, lo_inst=self.lo_inst)  # type: ignore

    def view_legend(self, visible: bool) -> None:
        """
        Shows or hides the legend.

        Args:
            visible (bool): ``True`` to show the legend, ``False`` to hide it.

        Note:
            If the legend is not found then it will be created if ``visible`` is ``True``.
        """
        legend = self.get_legend()
        if legend is not None:
            legend.show = visible
            return
        if visible:
            from .chart_legend import ChartLegend
            from ooo.dyn.drawing.line_style import LineStyle
            from ooo.dyn.drawing.fill_style import FillStyle

            legend = ChartLegend(owner=self, lo_inst=self.lo_inst)
            legend.set_property(LineStyle=LineStyle.NONE, FillStyle=FillStyle.SOLID, FillTransparence=100)
            self.set_legend(legend.component)

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document"""
        return self._owner
