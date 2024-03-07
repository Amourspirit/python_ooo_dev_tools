from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooo.dyn.chart.chart_axis_label_position import ChartAxisLabelPosition
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import chart2_axis_pos_label_position_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.axis.positioning.label_position_t import LabelPositionT
else:
    LoInst = Any
    LabelPositionT = Any


class Chart2AxisPosLabelPositionPartial:
    """
    Partial class for Chart Axis Position - Label Position.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_axis_pos_label_position",
            after_event="after_style_axis_pos_label_position",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_axis_pos_label_position(
        self, pos: ChartAxisLabelPosition = ChartAxisLabelPosition.NEAR_AXIS
    ) -> LabelPositionT | None:
        """
        Style Area Color.

        Args:
            pos (ChartAxisLabelPosition, optional): Specifies where to place the labels: near axis, near axis (other side), outside start, or outside end.
                Default is ``ChartAxisLabelPosition.NEAR_AXIS``.

        Raises:
            CancelEventError: If the event ``before_style_axis_pos_label_position`` is cancelled and not handled.

        Returns:
            LabelPositionT | None: Label Position instance or ``None`` if cancelled.

        Hint:
            - ``ChartAxisLabelPosition`` can be imported from ``ooo.dyn.chart.chart_axis_label_position``
        """
        return self.__styler.style(factory=chart2_axis_pos_label_position_factory, pos=pos)
