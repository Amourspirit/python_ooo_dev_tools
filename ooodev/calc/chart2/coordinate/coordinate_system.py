from __future__ import annotations
from typing import cast, TYPE_CHECKING, Tuple

from ooodev.calc.chart2.coordinate.coordinate_general import CoordinateGeneral

if TYPE_CHECKING:
    from com.sun.star.chart2 import XCoordinateSystem
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_diagram import ChartDiagram
    from ooodev.calc.chart2.chart_type import ChartType
    from ooodev.calc.chart2.chart_doc import ChartDoc


class CoordinateSystem(CoordinateGeneral):
    """Coordinate System Component."""

    def __init__(
        self, owner: ChartDiagram, chart_doc: ChartDoc, component: XCoordinateSystem, lo_inst: LoInst
    ) -> None:
        """
        Constructor

        Args:
            owner (ChartDiagram): Chart Diagram.
            component (XCoordinateSystem): UNO Chart2 ``XCoordinateSystem``.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        CoordinateGeneral.__init__(self, owner=owner, chart_doc=chart_doc, component=component, lo_inst=lo_inst)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.CoordinateSystem",)

    def get_chart_types(self) -> Tuple[ChartType[CoordinateGeneral], ...]:
        """
        Gets all chart types.

        Raises:
            ChartError: If an error occurs.

        Returns:
            Tuple[ChartType, ...]: A tuple of chart types.
        """
        return super().get_chart_types()  # type: ignore

    # endregion Overrides

    @property
    def component(self) -> ChartDiagram:
        """ChartDiagram General Component"""
        return cast("ChartDiagram", super().component)  # type: ignore

    @property
    def chart_type(self) -> ChartType[CoordinateSystem]:
        """Chart Type"""
        return super().chart_type  # type: ignore
