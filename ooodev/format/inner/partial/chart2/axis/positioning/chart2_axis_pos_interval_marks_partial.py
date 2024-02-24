from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.style_factory import chart2_axis_pos_interval_factory
from ooodev.loader import lo as mLo
from ooodev.format.inner.partial.factory_styler import FactoryStyler

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
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_axis_pos_interval_marks"
        self.__styler.before_event_name = "before_style_axis_pos_interval_marks"

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
        kwargs = {}
        if major is not None:
            kwargs["major"] = major
        if minor is not None:
            kwargs["minor"] = minor
        if pos is not None:
            kwargs["pos"] = pos
        return self.__styler.style(factory=chart2_axis_pos_interval_factory, **kwargs)
