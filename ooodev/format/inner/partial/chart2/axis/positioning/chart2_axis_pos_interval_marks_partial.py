from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_axis_pos_interval_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from ooo.dyn.chart.chart_axis_mark_position import ChartAxisMarkPosition
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.axis.positioning.interval_marks_t import IntervalMarksT
    from ooodev.format.inner.direct.chart2.axis.positioning.interval_marks import MarkKind
else:
    MarkKind = Any
    LoInst = Any
    ChartAxisMarkPosition = Any
    IntervalMarksT = Any


class Chart2AxisPosIntervalMarksPartial:
    """
    Partial class for Chart Axis Position Interval Marks.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_axis_pos_interval_marks(
        self, major: MarkKind | None = None, minor: MarkKind | None = None, pos: ChartAxisMarkPosition | None = None
    ) -> IntervalMarksT | None:
        """
        Style Area Color.

        Args:
            major (MarkKind, optional): Specifies the major tickmark style.
            minor (MarkKind, optional): Specifies the minor tickmark style.
            pos (ChartAxisMarkPosition, optional): Specifies where to place the marks: at labels, at axis, or at axis and labels.

        Raises:
            CancelEventError: If the event ``before_style_axis_pos_interval_marks`` is cancelled and not handled.

        Returns:
            IntervalMarksT | None: Interval Marks instance or ``None`` if cancelled.

        Hint:
            - ``MarkKind`` can be imported from ``ooodev.format.inner.direct.chart2.axis.positioning.interval_marks``
            - ``ChartAxisMarkPosition`` can be imported from ``ooo.dyn.chart.chart_axis_mark_position``
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_axis_pos_interval_marks.__qualname__)
            event_data: Dict[str, Any] = {
                "major": major,
                "minor": minor,
                "pos": pos,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_axis_pos_interval_marks", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_axis_pos_interval_marks")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Event has been cancelled.")
                    else:
                        return None
                else:
                    return None
            major = cargs.event_data.get("major", major)
            minor = cargs.event_data.get("minor", minor)
            pos = cargs.event_data.get("pos", pos)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_axis_pos_interval_factory(factory_name)
        fe = styler(major=major, minor=minor, pos=pos)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_axis_pos_interval_marks", EventArgs.from_args(cargs))  # type: ignore
        return fe
