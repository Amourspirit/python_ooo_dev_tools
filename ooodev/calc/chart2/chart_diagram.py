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
    from .chart_doc import ChartDoc
    from .coordinate.coordinate_general import CoordinateGeneral
    from ooodev.loader.inst.lo_inst import LoInst


class ChartDiagram(LoInstPropsPartial, DiagramComp, QiPartial, ServicePartial, StylePartial):
    """
    Class for managing Chart2 Diagram.
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
        DiagramComp.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        self._owner = owner

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

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document"""
        return self._owner
