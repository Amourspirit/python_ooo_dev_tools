from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.side import BorderLineKind
from ooodev.format.inner.direct.structs.side import LineSize
from ooodev.mock import mock_g
from ooodev.utils.color import StandardColor

if TYPE_CHECKING:
    from ooodev.format.inner.direct.structs.side import Side
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT


class WriteTableCellBordersPartial:
    """
    Partial class for Write Table Cell Borders.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_borders(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
    ) -> None:
        """
        Style Write Character Borders.

        Args:
            left (Side,, optional): Specifies the line style at the left edge.
            right (Side, optional): Specifies the line style at the right edge.
            top (Side, optional): Specifies the line style at the top edge.
            bottom (Side, optional): Specifies the line style at the bottom edge.
            border_side (Side, optional): Specifies the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored


        Raises:
            CancelEventError: If the event ``before_style_cell_borders`` is cancelled and not handled.

        Returns:
            Borders | None: Borders Style instance or ``None`` if cancelled.

        Hint:
            - ``BorderLine`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLine2`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLineKind`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``LineSize`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Side`` can be imported from ``ooodev.format.inner.direct.structs.side``

        """
        # pylint: disable=import-outside-toplevel

        def style_side(side: Side, comp: Any, prop_name: str, events: bool) -> Side | None:
            # pylint: disable=no-member
            bdr_side = side.copy()
            bdr_side.property_name = prop_name
            if events:
                bdr_side.add_event_observer(self.event_observer)  # type: ignore
            bdr_side.apply(comp)
            side.set_update_obj(comp)

        prop_keys = ("BottomBorder", "LeftBorder", "RightBorder", "TopBorder")

        comp = self.__component
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_borders.__qualname__)
            event_data: Dict[str, Any] = {
                "right": right,
                "left": left,
                "top": top,
                "bottom": bottom,
                "border_side": border_side,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_style_cell_borders", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_cell_borders")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            right = cargs.event_data.get("right", right)
            left = cargs.event_data.get("left", left)
            top = cargs.event_data.get("top", top)
            bottom = cargs.event_data.get("bottom", bottom)
            border_side = cargs.event_data.get("border_side", border_side)

            comp = cargs.event_data.get("this_component", comp)

        did_work = False
        if border_side is not None:
            did_work = True
            for prop_name in prop_keys:
                style_side(border_side, comp, prop_name, has_events)
        else:
            if right is not None:
                style_side(right, comp, prop_keys[2], has_events)
                did_work = True
            if left is not None:
                style_side(left, comp, prop_keys[1], has_events)
                did_work = True
            if top is not None:
                style_side(top, comp, prop_keys[3], has_events)
                did_work = True
            if bottom is not None:
                style_side(bottom, comp, prop_keys[0], has_events)
                did_work = True

        if did_work and cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            self.trigger_event("after_style_cell_borders", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore

    def style_borders_side(
        self,
        *,
        line: BorderLineKind = BorderLineKind.SOLID,
        color: Color = StandardColor.BLACK,
        width: LineSize | float | UnitT = LineSize.THIN,
    ) -> None:
        """
        Style All Write Character Borders.

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border. Default ``BorderLineKind.SOLID``.
            color (:py:data:`~.utils.color.Color`, optional): Color of the border. Default ``StandardColor.BLACK``
            width (LineSize, float, UnitT, optional): Contains the width in of a single line or the width of outer part of a double line (in ``pt`` units) or :ref:`proto_unit_obj`. If this value is zero, no line is drawn. Default ``LineSize.THIN``

        Raises:
            CancelEventError: If the event ``before_style_cell_borders`` is cancelled and not handled.

        Returns:
            Borders | None: Borders Style instance or ``None`` if cancelled.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
            - ``BorderLineKind`` can be imported from ``ooodev.format.inner.direct.structs.side``
            - ``LineSize`` can be imported from ``ooodev.format.inner.direct.structs.side``

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.structs.side import Side

        side = Side(line=line, color=color, width=width)
        return self.style_borders(border_side=side)


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.structs.side import Side
