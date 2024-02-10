from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno
from ooo.dyn.chart.chart_axis_label_position import ChartAxisLabelPosition
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_axis_pos_label_position_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

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
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

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
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_axis_pos_label_position.__qualname__)
            event_data: Dict[str, Any] = {
                "pos": pos,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_axis_pos_label_position", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_axis_pos_label_position")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style event has been cancelled.")
                    else:
                        return None
                else:
                    return None
            pos = cargs.event_data.get("pos", pos)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_axis_pos_label_position_factory(factory_name)
        fe = styler(pos=pos)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_axis_pos_label_position", EventArgs.from_args(cargs))  # type: ignore
        return fe
