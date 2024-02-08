from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
from ooodev.adapter.chart2.coordinate_system_partial import CoordinateSystemPartial
from ooodev.adapter.component_base import ComponentBase
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.adapter.chart2.chart_type_container_partial import ChartTypeContainerPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XCoordinateSystem
    from ooodev.loader.inst.lo_inst import LoInst
    from ..chart_diagram import ChartDiagram
    from ..chart_type import ChartType


class CoordinateGeneral(LoInstPropsPartial, ComponentBase, CoordinateSystemPartial, ChartTypeContainerPartial):
    """
    Class for managing Chart2 Coordinate General.
    """

    # pylint: disable=unused-argument

    def __init__(self, owner: ChartDiagram, component: XCoordinateSystem, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (ChartDiagram): Chart Diagram.
            component (Any): UNO Chart2 Coordinate General.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ComponentBase.__init__(self, component)
        CoordinateSystemPartial.__init__(self, component=component)
        ChartTypeContainerPartial.__init__(self, component=component, interface=None)  # type: ignore
        self.__owner = owner
        self.__chart_type = None

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region ChartTypeContainerPartial overrides
    def get_chart_types(self) -> Tuple[ChartType[CoordinateGeneral], ...]:
        """
        Gets all chart types.

        Raises:
            ChartError: If an error occurs.

        Returns:
            Tuple[ChartType, ...]: A tuple of chart types.
        """
        try:
            from ..chart_type import ChartType

            c_types = super().get_chart_types()
            if not c_types:
                return ()
            return tuple(ChartType(owner=self, component=c_type, lo_inst=self.lo_inst) for c_type in c_types)
        except Exception as e:
            raise mEx.ChartError("Error getting chart types") from e

    # endregion ChartTypeContainerPartial overrides

    # region Properties
    @property
    def component(self) -> XCoordinateSystem:
        """Coordinate General Component"""
        return self._ComponentBase__get_component()  # type: ignore

    @property
    def chart_diagram(self) -> ChartDiagram:
        """Chart Diagram"""
        return self.__owner

    @property
    def chart_type(self) -> ChartType[CoordinateGeneral]:
        """Chart Type"""
        if self.__chart_type is None:
            try:
                self.__chart_type = self.get_chart_types()[0]
            except mEx.ChartError:
                raise
            except Exception as e:
                raise mEx.ChartError("Error getting chart type") from e
        return self.__chart_type

    # endregion Properties
