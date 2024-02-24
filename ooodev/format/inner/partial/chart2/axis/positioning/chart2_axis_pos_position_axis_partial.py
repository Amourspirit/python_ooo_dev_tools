from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import chart2_axis_pos_position_axis_factory
from ooodev.loader import lo as mLo

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
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_area_gradient"
        self.__styler.before_event_name = "before_style_area_gradient"

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
