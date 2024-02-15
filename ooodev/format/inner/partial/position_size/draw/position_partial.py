from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import draw_position_size_position_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.draw.position_size.position_t import PositionT
    from ooodev.units import UnitT
else:
    PositionT = Any
    UnitT = Any


class PositionPartial:
    """
    Partial class for Position.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_position(
        self, x: float | UnitT, y: float | UnitT, base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT
    ) -> PositionT | None:
        """
        Style Area Color.

        Args:
            x (float, UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            y (float, UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind, optional): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Raises:
            CancelEventError: If the event ``before_style_position`` is cancelled and not handled.

        Returns:
            PositionT | None: Position instance or ``None`` if cancelled.

        Hint:
            - ``ShapeBasePointKind`` can be imported from ``ooodev.utils.kind.shape_base_point_kind``
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_position.__qualname__)
            event_data: Dict[str, Any] = {
                "x": x,
                "y": y,
                "base_point": base_point,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_position", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_position")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                else:
                    return None
            x = cargs.event_data.get("x", x)
            y = cargs.event_data.get("y", y)
            base_point = cargs.event_data.get("base_point", base_point)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = draw_position_size_position_factory(factory_name)
        fe = styler(pos_x=x, pos_y=y, base_point=base_point)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_position", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_position_get(self) -> PositionT | None:
        """
        Gets the Position Style.

        Raises:
            CancelEventError: If the event ``before_style_position_get`` is cancelled and not handled.

        Returns:
            PositionT | None: Position style or ``None`` if cancelled.
        """
        comp = self.__component
        factory_name = self.__factory_name
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_position_get.__qualname__)
            event_data: Dict[str, Any] = {
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_position_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_position_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = draw_position_size_position_factory(factory_name)
        try:
            style = styler.from_obj(comp)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style
