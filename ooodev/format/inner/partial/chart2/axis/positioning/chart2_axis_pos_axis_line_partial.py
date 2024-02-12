from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_axis_pos_line_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooo.dyn.chart.chart_axis_position import ChartAxisPosition

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.axis.positioning.axis_line_t import AxisLineT
    from ooodev.utils.gen_util import NULL_OBJ, TNullObj
else:
    AxisLineT = Any
    NULL_OBJ = Any
    TNullObj = Any


class Chart2AxisPosAxisLinePartial:
    """
    Partial class for Chart Axis Position Axis Line.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

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
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_axis_pos_axis_line.__qualname__)
            event_data: Dict[str, Any] = {
                "cross": cross,
                "value": value,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_axis_pos_axis_line", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_axis_pos_axis_line")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style event has been cancelled.")
                    else:
                        return None
                else:
                    return None
            cross = cargs.event_data.get("cross", cross)
            value = cargs.event_data.get("value", value)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_axis_pos_line_factory(factory_name)
        fe = styler(cross=cross, value=value)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_axis_pos_axis_line", EventArgs.from_args(cargs))  # type: ignore
        return fe
