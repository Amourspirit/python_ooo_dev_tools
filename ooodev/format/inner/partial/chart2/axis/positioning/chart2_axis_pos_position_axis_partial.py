from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_axis_pos_position_axis_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

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
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

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
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_axis_pos_position_axis.__qualname__)
            event_data: Dict[str, Any] = {
                "on_mark": on_mark,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_axis_pos_position_axis", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_axis_pos_position_axis")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style event has been cancelled.")
                else:
                    return None
            on_mark = cargs.event_data.get("on_mark", on_mark)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_axis_pos_position_axis_factory(factory_name)
        fe = styler(on_mark=on_mark)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_axis_pos_position_axis", EventArgs.from_args(cargs))  # type: ignore
        return fe
