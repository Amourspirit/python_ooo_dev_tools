from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from ooodev.mock import mock_g
from ooodev.adapter.chart2.diagram_comp import DiagramComp
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.calc.chart2.kind.chart_title_kind import ChartTitleKind
from ooodev.calc.chart2.kind.chart_diagram_kind import ChartDiagramKind
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_doc import ChartDoc
    from ooodev.calc.chart2.coordinate.coordinate_general import CoordinateGeneral
    from ooodev.calc.chart2.chart_title import ChartTitle
    from ooodev.calc.chart2.chart_legend import ChartLegend
    from ooodev.calc.chart2.chart_wall import ChartWall
    from ooodev.calc.chart2.chart_floor import ChartFloor


class ChartDiagram(
    LoInstPropsPartial,
    DiagramComp,
    ChartDocPropPartial,
    QiPartial,
    ServicePartial,
    StylePartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
):
    """
    Class for managing Chart2 Diagram.
    """

    def __init__(
        self, owner: ChartDoc, component: Any, diagram_kind: ChartDiagramKind, lo_inst: LoInst | None = None
    ) -> None:
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
        ChartDocPropPartial.__init__(self, chart_doc=owner)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        self._wall = None
        self._floor = None
        self._diagram_kind = diagram_kind

    def get_title(self) -> ChartTitle[ChartDiagram] | None:
        """Gets the Title Diagram Component. This might be considered to be a subtitle."""
        # pylint: disable=import-outside-toplevel
        from .chart_title import ChartTitle

        comp = self.get_title_object()
        if comp is None:
            return None
        if self.diagram_kind == ChartDiagramKind.FIRST:
            chart_title_kind = ChartTitleKind.SUBTITLE
        else:
            chart_title_kind = ChartTitleKind.UNKNOWN
        return ChartTitle(
            owner=self, chart_doc=self.chart_doc, component=comp, title_kind=chart_title_kind, lo_inst=self.lo_inst
        )

    def set_title(self, title: str) -> ChartTitle[ChartDiagram]:
        """Adds a Chart Title."""
        # pylint: disable=import-outside-toplevel
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
        if self.diagram_kind == ChartDiagramKind.FIRST:
            chart_title_kind = ChartTitleKind.SUBTITLE
        else:
            chart_title_kind = ChartTitleKind.UNKNOWN

        return ChartTitle(
            owner=self,
            chart_doc=self.chart_doc,
            component=titled.getTitleObject(),
            title_kind=chart_title_kind,
            lo_inst=self.lo_inst,
        )

    def get_coordinate_system(self) -> CoordinateGeneral | None:
        """Gets the first Coordinate System Component."""
        # sourcery skip: lift-return-into-if

        coord_sys = super().get_coordinate_systems()
        if not coord_sys:
            return None
        first = coord_sys[0]
        # pylint: disable=import-outside-toplevel
        if mInfo.Info.support_service(first, "com.sun.star.chart2.CoordinateSystem"):
            from .coordinate.coordinate_system import CoordinateSystem

            result = CoordinateSystem(owner=self, chart_doc=self.chart_doc, component=first, lo_inst=self.lo_inst)
        else:
            from .coordinate.coordinate_general import CoordinateGeneral

            result = CoordinateGeneral(owner=self, chart_doc=self.chart_doc, component=first, lo_inst=self.lo_inst)

        return result

    # region CoordinateSystemContainerPartial overrides
    def get_coordinate_systems(self) -> Tuple[CoordinateGeneral, ...]:
        """
        Gets all coordinate systems
        """
        # pylint: disable=import-outside-toplevel
        from .coordinate.coordinate_system import CoordinateSystem
        from .coordinate.coordinate_general import CoordinateGeneral

        result = []
        coord_sys = super().get_coordinate_systems()
        for sys in coord_sys:
            if mInfo.Info.support_service(sys, "com.sun.star.chart2.CoordinateSystem"):
                result.append(
                    CoordinateSystem(
                        owner=self, chart_doc=self.chart_doc, component=coord_sys[0], lo_inst=self.lo_inst
                    )
                )
            else:
                result.append(
                    CoordinateGeneral(
                        owner=self, chart_doc=self.chart_doc, component=coord_sys[0], lo_inst=self.lo_inst
                    )
                )
        return tuple(result)

    # endregion CoordinateSystemContainerPartial overrides

    def view_legend(self, visible: bool) -> None:
        """
        Shows or hides the legend.

        Args:
            visible (bool): ``True`` to show the legend, ``False`` to hide it.

        Note:
            If the legend is not found then it will be created if ``visible`` is ``True``.
        """
        # pylint: disable=import-outside-toplevel
        legend = self.get_legend()
        if legend is not None:
            legend.show = visible
            return
        if visible:
            from ooo.dyn.drawing.line_style import LineStyle
            from ooo.dyn.drawing.fill_style import FillStyle
            from .chart_legend import ChartLegend

            legend = ChartLegend(owner=self, chart_doc=self.chart_doc, lo_inst=self.lo_inst)
            legend.set_property(LineStyle=LineStyle.NONE, FillStyle=FillStyle.SOLID, FillTransparence=100)
            self.set_legend(legend.component)

    # region DiagramPartial (XDiagram) Overrides
    def get_legend(self) -> ChartLegend | None:
        """
        Gets the Legend Component.

        Returns:
            ChartLegend | None: Legend Component if found, otherwise ``None``.
        """
        # pylint: disable=import-outside-toplevel
        legend = self.component.getLegend()
        if legend is None:
            return None
        from .chart_legend import ChartLegend

        return ChartLegend(owner=self, chart_doc=self.chart_doc, component=legend, lo_inst=self.lo_inst)  # type: ignore

    def get_wall(self) -> ChartWall:
        """
        Gets the ``ChartWall`` that contains the property set that determines the visual appearance of the wall.

        Returns:
            ChartWall: Wall Component.
        """
        return self.wall

    def get_floor(self) -> ChartFloor:
        """
        Gets the ``ChartFloor`` that contains the property set that determines the visual appearance of the wall.

        Returns:
            ChartFloor: Floor Component.
        """
        return self.floor

    # endregion DiagramPartial (XDiagram) Overrides

    @property
    def wall(self) -> ChartWall:
        """Gets the ``ChartFloor`` that contains the property set that determines the visual appearance of the wall."""
        # pylint: disable=import-outside-toplevel
        if self._wall is None:
            from .chart_wall import ChartWall

            self._wall = ChartWall(owner=self, component=self.component.getWall(), lo_inst=self.lo_inst)
        return self._wall

    @property
    def floor(self) -> ChartFloor:
        """Gets the ``ChartFloor`` that contains the property set that determines the visual appearance of the wall."""
        # pylint: disable=import-outside-toplevel
        if self._floor is None:
            from .chart_floor import ChartFloor

            self._floor = ChartFloor(owner=self, component=self.component.getFloor(), lo_inst=self.lo_inst)
        return self._floor

    @property
    def diagram_kind(self) -> ChartDiagramKind:
        """Gets the diagram kind."""
        return self._diagram_kind


if mock_g.FULL_IMPORT:
    from com.sun.star.chart2 import XFormattedString
    from com.sun.star.chart2 import XTitle
    from com.sun.star.chart2 import XTitled
    from ooo.dyn.drawing.fill_style import FillStyle
    from ooo.dyn.drawing.line_style import LineStyle
    from .chart_floor import ChartFloor
    from .chart_legend import ChartLegend
    from .chart_title import ChartTitle
    from .chart_wall import ChartWall
    from .coordinate.coordinate_general import CoordinateGeneral
    from .coordinate.coordinate_system import CoordinateSystem
