from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooo.dyn.chart.chart_axis_position import ChartAxisPosition
from ooodev.format.inner.style_factory import chart2_axis_pos_line_factory
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.gen_util import NULL_OBJ

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.axis.positioning.axis_line_t import AxisLineT
    from ooodev.utils.gen_util import TNullObj
else:
    LoInst = Any
    AxisLineT = Any
    TNullObj = Any


class Chart2AxisPosAxisLinePartial:
    """
    Partial class for Chart Axis Position Axis Line.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_axis_pos_axis_line",
            after_event="after_style_axis_pos_axis_line",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_axis_pos_axis_line(
        self,
        cross: ChartAxisPosition | None = None,
        value: float | None | TNullObj = NULL_OBJ,
    ) -> AxisLineT | None:
        """
        Style Area Color.

        Args:
            cross(ChartAxisPosition, optional): The position where the axis crosses the other axis.
            value (float, None, optional): The value where the axis crosses the other axis.

        Raises:
            CancelEventError: If the event ``before_style_axis_pos_axis_line`` is cancelled and not handled.

        Returns:
            AxisLineT | None: Axis Line instance or ``None`` if cancelled.

        Hint:
            - ``ChartAxisPosition`` can be imported from ``ooo.dyn.chart.chart_axis_position``
        """
        factory = chart2_axis_pos_line_factory
        return self.__styler.style(factory=factory, cross=cross, value=value)
