from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.style_factory import chart2_axis_pos_interval_factory
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_axis_pos_interval_marks",
            after_event="after_style_axis_pos_interval_marks",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

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
