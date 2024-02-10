from __future__ import annotations
from typing import Any, TYPE_CHECKING, Callable

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.format.inner.partial.factory_name_base import FactoryNameBase

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.callable_any_t import CallableAnyT
else:
    LoInst = Any
    CallableAnyT = Any


class FactoryStyler(FactoryNameBase):
    """
    Class for Line Properties.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        super().__init__(factory_name, component, lo_inst)
        self.before_event_name = "before_style_border_line"
        self.after_event_name = "after_style_border_line"

    def style(self, factory: Callable[[str], Any], **kwargs) -> Any:
        """
        Style Font.

        Raises:
            CancelEventError: If the ``before_*`` event is cancelled and not handled.

        Returns:
            Any: Style instance or ``None`` if cancelled.
        """
        comp = self._component
        factory_name = self._factory_name
        has_events = False
        cargs = None
        event_data = kwargs.copy()
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style.__qualname__)
            event_data["factory_name"] = factory_name
            event_data["this_component"] = comp

            cargs.event_data = event_data
            self.trigger_event(self.before_event_name, cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", self.before_event_name)
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                    else:
                        return None
                else:
                    return None
            comp = cargs.event_data.pop("this_component", comp)
            factory_name = cargs.event_data.pop("factory_name", factory_name)
            event_data = cargs.event_data

        styler = factory(factory_name)
        fe = styler(**event_data)
        # fe.factory_name = factory_name

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self._lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event(self.after_event_name, EventArgs.from_args(cargs))  # type: ignore
        return fe
