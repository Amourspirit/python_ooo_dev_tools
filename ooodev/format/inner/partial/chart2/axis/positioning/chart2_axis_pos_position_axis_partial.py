from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import chart2_axis_pos_position_axis_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.axis.positioning.position_axis_t import PositionAxisT
else:
    LoInst = Any
    PositionAxisT = Any


class Chart2AxisPosPositionAxisPartial:
    """
    Partial class for Chart Axis Position - Position Axis.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_gradient",
            after_event="after_style_area_gradient",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_axis_pos_position_axis(self, on_mark: bool = True) -> PositionAxisT | None:
        """
        Style Area Color.

        Args:
            on_mark(bool, optional): Specifies that the axis is position.
                If ``True``, specifies that the axis is positioned on the first/last tickmarks. This makes the data points visual representation begin/end at the value axis.
                If ``False``, specifies that the axis is positioned between the tickmarks. This makes the data points visual representation begin/end at a distance from the value axis.
                Default is ``True``.

        Raises:
            CancelEventError: If the event ``before_style_axis_pos_position_axis`` is cancelled and not handled.

        Returns:
            PositionAxisT | None: Position Axis instance or ``None`` if cancelled.
        """
        return self.__styler.style(factory=chart2_axis_pos_position_axis_factory, on_mark=on_mark)
