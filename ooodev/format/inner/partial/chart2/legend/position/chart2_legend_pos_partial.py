from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.mock import mock_g
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.exceptions import ex as mEx
from ooodev.events.style_named_event import StyleNameEvent

from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING:  # or DOCS_BUILDING:
    from ooodev.format.inner.direct.chart2.legend.position.position import Position
    from ooo.dyn.chart2.legend_position import LegendPosition
    from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
else:
    LegendPosition = Any
    DirectionModeKind = Any


class Chart2LegendPosPartial:
    """
    Partial class for Chart2 Legend Position.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_position(
        self,
        *,
        pos: LegendPosition | None = None,
        mode: DirectionModeKind | None = None,
        no_overlap: bool | None = None,
    ) -> Position | None:
        """
        Style Area Color.

        Args:
            pos (LegendPosition | None, optional): Specifies the position of the legend.
            mode (DirectionModeKind, optional): Specifies the writing direction.
            no_overlap (bool | None, optional): Show the legend without overlapping the chart.

        Raises:
            CancelEventError: If the event ``before_style_chart2_legend_pos`` is cancelled and not handled.

        Returns:
            PositionT | None: Chart Legend Position style instance or ``None`` if cancelled.

        Hint:
            - ``LegendPosition`` can be imported from ``ooo.dyn.chart2.legend_position``.
            - ``DirectionModeKind`` can be imported from ``ooodev.format.inner.direct.chart2.title.alignment.direction``.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.chart2.legend.position.position import Position

        comp = self.__component
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_position.__qualname__)
            event_data: Dict[str, Any] = {
                "pos": pos,
                "mode": mode,
                "no_overlap": no_overlap,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_style_chart2_legend_pos", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_chart2_legend_pos")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Area Gradient has been cancelled.")
                else:
                    return None
            pos = cargs.event_data.get("pos", pos)
            mode = cargs.event_data.get("mode", mode)
            no_overlap = cargs.event_data.get("no_overlap", no_overlap)
            comp = cargs.event_data.get("this_component", comp)

        fe = Position(
            pos=pos,
            mode=mode,
            no_overlap=no_overlap,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        fe.apply(comp)
        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event("after_style_chart2_legend_pos", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.chart2.legend.position.position import Position
