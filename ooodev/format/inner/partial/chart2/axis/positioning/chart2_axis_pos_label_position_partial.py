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
from ooodev.format.inner.partial.factory_styler import FactoryStyler

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
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_axis_pos_label_position"
        self.__styler.before_event_name = "before_style_axis_pos_label_position"

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
