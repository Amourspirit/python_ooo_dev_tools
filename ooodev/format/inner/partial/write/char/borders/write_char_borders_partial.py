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
from ooodev.loader import lo as mLo
from ooodev.mock import mock_g
from ooodev.utils.color import StandardColor
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.format.inner.direct.write.char.border.padding import Padding
    from ooodev.format.inner.direct.write.char.border.shadow import Shadow
    from ooodev.format.inner.direct.structs.side import Side
    from ooodev.format.inner.direct.write.char.border.borders import Borders
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT


class WriteCharBordersPartial:
    """
    Partial class for Write Char Borders.
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
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> Borders | None:
        """
        Style Write Character Borders.

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (~ooodev.format.inner.direct.write.char.border.shadow.Shadow | None, optional): Character Shadow
            padding (~ooodev.format.inner.direct.write.char.border.padding.Padding | None, optional): Character padding

        Raises:
            CancelEventError: If the event ``before_style_char_borders`` is cancelled and not handled.

        Returns:
            Borders | None: Attribute Options Style instance or ``None`` if cancelled.

        Hint:
            - ``BorderLine`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLine2`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLineKind`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Borders`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``LineSize`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Padding`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Shadow`` can be imported from ``ooodev.format.inner.direct.write.char.border.shadow``
            - ``ShadowFormat`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Side`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``side`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Sides`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``ShadowLocation`` can be imported ``from ooo.dyn.table.shadow_location``

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.char.border.borders import Borders

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
                "shadow": shadow,
                "padding": padding,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_style_char_borders", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_char_borders")
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
            shadow = cargs.event_data.get("shadow", shadow)
            padding = cargs.event_data.get("padding", padding)

            comp = cargs.event_data.get("this_component", comp)

        fe = Borders(
            right=right,
            left=left,
            top=top,
            bottom=bottom,
            all=border_side,
            shadow=shadow,
            padding=padding,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        if isinstance(self, LoInstPropsPartial):
            lo_inst = self.lo_inst
        else:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst):
            fe.apply(comp)

        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event("after_style_char_borders", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe

    def style_borders_side(
        self,
        *,
        line: BorderLineKind = BorderLineKind.SOLID,
        color: Color = StandardColor.BLACK,
        width: LineSize | float | UnitT = LineSize.THIN,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> Borders | None:
        """
        Style All Write Character Borders.

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border. Default ``BorderLineKind.SOLID``.
            color (:py:data:`~.utils.color.Color`, optional): Color of the border. Default ``StandardColor.BLACK``
            width (LineSize, float, UnitT, optional): Contains the width in of a single line or the width of outer part of a double line (in ``pt`` units) or :ref:`proto_unit_obj`. If this value is zero, no line is drawn. Default ``LineSize.THIN``
            shadow (Shadow | None, optional): Character Shadow
            padding (Padding | None, optional): Character padding

        Raises:
            CancelEventError: If the event ``before_style_char_borders`` is cancelled and not handled.

        Returns:
            Borders | None: Borders Style instance or ``None`` if cancelled.

        Hint:
            - ``BorderLine`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLine2`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``BorderLineKind`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``LineSize`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Padding`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Shadow`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``ShadowFormat`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``ShadowLocation`` can be imported ``from ooo.dyn.table.shadow_location``

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.structs.side import Side

        side = Side(line=line, color=color, width=width)
        return self.style_borders(
            border_side=side,
            shadow=shadow,
            padding=padding,
        )

    def style_borders_padding(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
        all_sides: float | UnitT | None = None,
    ) -> Borders | None:
        """
        Style Padding for Write Characters.

        Args:
            left (float, UnitT, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitT, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitT, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitT,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all_sides (float, UnitT, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            CancelEventError: If the event ``before_style_char_borders`` is cancelled and not handled.

        Returns:
            Borders | None: Borders Style instance or ``None`` if cancelled.

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.char.border.padding import Padding

        padding = Padding(left=left, right=right, top=top, bottom=bottom, all=all_sides)
        return self.style_borders(padding=padding)


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.write.char.border.padding import Padding
    from ooodev.format.inner.direct.structs.side import Side
    from ooodev.format.inner.direct.write.char.border.borders import Borders
