from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import draw_position_size_protect_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.draw.position_size.protect_t import ProtectT
else:
    ProtectT = Any


class ProtectPartial:
    """
    Partial class for Size.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_protect(self, position: bool | None = None, size: bool | None = None) -> ProtectT | None:
        """
        Style Area Color.

        Args:
            position (bool, optional): Specifies position protection.
            size (bool, optional): Specifies size protection.

        Raises:
            CancelEventError: If the event ``before_style_protect`` is cancelled and not handled.

        Returns:
            ProtectT | None: Position instance or ``None`` if cancelled.
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_protect.__qualname__)
            event_data: Dict[str, Any] = {
                "position": position,
                "size": size,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_protect", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_protect")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                else:
                    return None
            position = cargs.event_data.get("position", position)
            size = cargs.event_data.get("size", size)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = draw_position_size_protect_factory(factory_name)
        fe = styler(position=position, size=size)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_protect", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_protect_get(self) -> ProtectT | None:
        """
        Gets the Protect Style.

        Raises:
            CancelEventError: If the event ``before_style_protect_get`` is cancelled and not handled.

        Returns:
            ProtectT | None: Protect style or ``None`` if cancelled.
        """
        comp = self.__component
        factory_name = self.__factory_name
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_protect_get.__qualname__)
            event_data: Dict[str, Any] = {
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_protect_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_protect_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = draw_position_size_protect_factory(factory_name)
        try:
            style = styler.from_obj(comp)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style
